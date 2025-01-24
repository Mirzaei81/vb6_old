VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmUsePercent 
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14340
   Icon            =   "frmUsePercent.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   14340
   Begin VB.ListBox lstGoodLevel2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   11535
      Style           =   1  'Checkbox
      TabIndex        =   16
      Top             =   7875
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.ListBox lstGoodLevel1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   11505
      Style           =   1  'Checkbox
      TabIndex        =   14
      Top             =   5265
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2520
      Width           =   2895
      Begin VB.CommandButton cmdCalAvgBuyPrice 
         BackColor       =   &H000000C0&
         Caption         =   "ãÍÇÓÈå ãíÇäíä ÎÑíÏ"
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
         TabIndex        =   12
         Top             =   1200
         Width           =   2655
      End
      Begin MSMask.MaskEdBox txtDateTo 
         Height          =   465
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   820
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtDateFrom 
         Height          =   465
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   820
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÊÇ ÊÇÑíÎ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇÒ ÊÇÑíÎ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2895
      Begin VB.CommandButton SetUsedPercent 
         BackColor       =   &H000000C0&
         Caption         =   "Èå ÑæÒ ÑÓÇäí ÈåÇÁ ÊãÇã ÔÏå ßÇáÇí ÂãÇÏå ÝÑæÔ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   2655
      End
      Begin FLWCtrls.FWProgressBar FWProgressBar1 
         Height          =   375
         Left            =   120
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         Max             =   10
         BorderStyle     =   10
      End
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   12720
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
   Begin VSFlex7LCtl.VSFlexGrid vsGoodFirst 
      Height          =   5145
      Left            =   120
      TabIndex        =   0
      Top             =   4995
      Width           =   14055
      _cx             =   24791
      _cy             =   9075
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
      BackColorFixed  =   16777152
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
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
      AutoResize      =   -1  'True
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
   Begin VSFlex7LCtl.VSFlexGrid vsGood 
      Height          =   3945
      Left            =   3120
      TabIndex        =   3
      Top             =   960
      Width           =   11115
      _cx             =   19606
      _cy             =   6959
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
      BackColorFixed  =   16761024
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
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   13
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
      AutoResize      =   -1  'True
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmUsePercent.frx":A4C2
      TabIndex        =   11
      Top             =   0
      Width           =   480
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   480
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
   Begin VB.Label lblGoodLevel2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ñæå ÝÑÚí ˜ÇáÇåÇ"
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
      Left            =   11415
      TabIndex        =   17
      Top             =   7515
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Label lblGoodLevel1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ñæå ÇÕáí ˜ÇáÇåÇ"
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
      Left            =   11490
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÚííä ÖÑíÈ ãÕÑÝ æ ãÍÇÓÈå ÞíãÊ ÊãÇã ÔÏå ßÇáÇí ÂãÇÏå ÝÑæÔ"
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
      Height          =   495
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   7935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "˜ÇáÇ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   12840
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ãæÇÏ Çæáíå"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   525
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4560
      Width           =   1935
   End
End
Attribute VB_Name = "frmUsePercent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter
Dim clsDate As New clsDate
Dim CodeVsGood As String
Dim CodeVsGoodFirst As String
Public Sub ChangeLanguage()
   
   
   FillvsGood
   FillVsGoodFirst
   
End Sub

Public Sub Cancel()
    
    Me.SetFocus
    UpdateVsGoodFirst
    
    MyFormAddEditMode = ViewMode 'View Mode
    SetFirstToolBar
    
End Sub

Public Sub Edit()
'    If vsGood.Row > 1 Then
'        vsGood.TextMatrix(vsGood.Row, 4) = -1
    
        MyFormAddEditMode = EditMode 'Edit Mode
        SetFirstToolBar
'    End If
 CodeVsGood = "0"
 CodeVsGoodFirst = ""
 
End Sub

Public Sub ExitForm()
    Unload Me
    
End Sub

Public Sub FillvsGood()

    Dim i As Integer
    Dim Rst As New ADODB.Recordset
    
   
    With vsGood
        .Rows = 1
    
    ReDim Parameter(1) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SupportedGoodType", adInteger, 4, EnumGoodType.forSale)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Levels_Supported", Parameter)
        
        If Rst.EOF = True And Rst.BOF = True Then
            Set Rst = Nothing
            Exit Sub
        End If
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst.Fields("Code").Value
            .TextMatrix(i, 2) = Rst.Fields("Level1").Value
            .TextMatrix(i, 3) = Rst.Fields("Level2").Value
            .TextMatrix(i, 4) = 0
            .TextMatrix(i, 5) = Rst.Fields("Deslevel1").Value
            .TextMatrix(i, 6) = Rst.Fields("Deslevel2").Value
            .TextMatrix(i, 7) = Left(Rst.Fields("Name").Value, 25)
            .TextMatrix(i, 8) = Rst.Fields("FinalPrice").Value
            .TextMatrix(i, 9) = IIf(IsNull(Rst!ChargeCooking), "", Rst!ChargeCooking)
            .TextMatrix(i, 10) = IIf(IsNull(Rst!ChargeServe), "", Rst!ChargeServe)
            .TextMatrix(i, 11) = IIf(IsNull(Rst!PercentOverFlow), "", Rst!PercentOverFlow)
            
            Rst.MoveNext
        Wend
                
        .Row = 0
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
    Set Rst = Nothing
End Sub

Public Sub FillVsGoodFirst()

    If lstGoodLevel1.ListCount < 1 Then Exit Sub
    If lstGoodLevel2.ListCount < 1 Then Exit Sub

    Dim Rst As New ADODB.Recordset
    
    MyFormAddEditMode = ViewMode 'VIEW Mode
    SetFirstToolBar
    
    vsGoodFirst.Rows = 1

    
    Dim i As Integer
    Dim j As Integer
    Dim intSelectedLevel1 As Integer
    Dim intSelectedLevel2 As Integer
    Dim strSelectedLevels As String

    
    intSelectedLevel1 = -1
    intSelectedLevel2 = -1
    
    For i = 0 To lstGoodLevel1.ListCount - 1
        If lstGoodLevel1.Selected(i) = True Then
            intSelectedLevel1 = lstGoodLevel1.ItemData(i)
        End If
    Next i
    
    strSelectedLevels = ""
    For i = 0 To lstGoodLevel2.ListCount - 1
        If lstGoodLevel2.Selected(i) = True Then
            intSelectedLevel2 = i
            strSelectedLevels = strSelectedLevels + "," + CStr(lstGoodLevel2.ItemData(i))
        End If
    Next i
    
    With vsGoodFirst
        .Rows = 1
         ReDim Parameter(1) As Parameter
          
         Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
         Parameter(1) = GenerateInputParameter("@NotSupportedGoodType", adInteger, 4, EnumGoodType.forSale)
         
         Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Levels", Parameter)
        If Rst.EOF = True And Rst.BOF = True Then
            Set Rst = Nothing
            Exit Sub
        End If
'        Rst.moveFirst
        While Rst.EOF <> True
           If (Rst.Fields("Level1").Value = intSelectedLevel1 Or intSelectedLevel1 = -1) And (InStr(1, strSelectedLevels, Rst.Fields("Level2").Value, vbBinaryCompare) > 0 Or intSelectedLevel2 = -1) Then
               .Rows = .Rows + 1
               i = .Rows - 1
               .TextMatrix(i, 0) = i
               .TextMatrix(i, 1) = Rst.Fields("code").Value
               .TextMatrix(i, 2) = Rst.Fields("Level1").Value
               .TextMatrix(i, 3) = Rst.Fields("Level2").Value
               .TextMatrix(i, 4) = 0
               .TextMatrix(i, 5) = Rst.Fields("Deslevel1").Value
               .TextMatrix(i, 6) = Rst.Fields("Deslevel2").Value
               .TextMatrix(i, 7) = Rst.Fields("Name").Value
               .TextMatrix(i, 8) = Rst.Fields("UnitDescription").Value
          End If
            
          Rst.MoveNext
        Wend
        
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
    Set Rst = Nothing

End Sub

Public Sub Update()

    Dim i As Integer
    Dim j As Integer
    Dim strSelectedGood As String
    Dim ChargeCooking As Integer
    Dim ChargeServe As Integer
    Dim PercentOverFlow As Integer
    Dim s As String
    Dim T As String
    Dim U As String
    Dim p As String
    
    
    vsGoodFirst_ValidateEdit vsGoodFirst.Row, vsGoodFirst.Col, False
    
    Select Case MyFormAddEditMode
    
        Case EditMode 'Edit
            If CodeVsGood <> "0" And CodeVsGood <> "-1" Then
                With vsGood
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 1) = CodeVsGood Then
                        If Val(.TextMatrix(i, 11)) >= 0 And Val(.TextMatrix(i, 11)) <= 100 Then
                            ChargeCooking = Val(.TextMatrix(i, 9))
                            ChargeServe = Val(.TextMatrix(i, 10))
                            PercentOverFlow = Val(.TextMatrix(i, 11))
                            Exit For
                        Else
                            ShowDisMessage "ÏÑÕÏ ÓÑÈÇÑ ÈÇíÏ ˜ãÊÑ ÇÒ 100 ÈÇÔÏ", 2000
                            .ShowCell i, 11: .Select i, 11: .EditCell
                        End If
                    End If
                Next i
                End With
                ReDim Parameter(3) As Parameter
                Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, CodeVsGood)
                Parameter(1) = GenerateInputParameter("@ChargeCooking", adInteger, 4, ChargeCooking)
                Parameter(2) = GenerateInputParameter("@ChargeServe", adInteger, 4, ChargeServe)
                Parameter(3) = GenerateInputParameter("@PercentOverFlow", adInteger, 4, PercentOverFlow)
                
                RunParametricStoredProcedure "Update_tblTotal_ChargeGood", Parameter
            End If
            With vsGood
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 4) = -1 Then
                        strSelectedGood = .TextMatrix(i, 1)
                        Exit For
                    End If
                Next i
            End With
            
            If strSelectedGood = "" And (CodeVsGood = "0" Or CodeVsGood = "-1") Then
            
                    frmMsg.fwlblMsg.Caption = "áØÝÇ ÇÈÊÏÇ í˜ ˜ÇáÇ ÇäÊÎÇÈ äãÇííÏ "
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ÞÈæá"
                    frmMsg.Show vbModal
                    
                    Exit Sub
            ElseIf strSelectedGood = "" And CodeVsGood <> "0" And CodeVsGood <> "-1" Then
                frmMsg.fwlblMsg.Caption = "ÊÛííÑÇÊ åÒíäå ÇäÌÇã ÔÏ "
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ÞÈæá"
                    frmMsg.Show vbModal
                    MyFormAddEditMode = ViewMode 'View Mode
                    SetFirstToolBar
                    Exit Sub
            
            End If
            
            
            
            
            ReDim Parameter(4) As Parameter
            
            With vsGoodFirst
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 4) = -1 Then
                        For j = 9 To .Cols - 1 Step 2
                            If Val(.TextMatrix(i, j)) > 0 Then
                                
                                s = s & .TextMatrix(i, 1) & ","
                                
                                T = T & .Cell(flexcpData, 0, j, 0, j) & ","
                                
                                U = U & .TextMatrix(i, j) & ","
                                
                                p = p & .TextMatrix(i, j + 1) & ","
                                
                            End If
                        Next j
                    End If
                Next i
                
                If s <> "" Then
                    s = Left(s, Len(s) - 1)
                    T = Left(T, Len(T) - 1)
                    U = Left(U, Len(U) - 1)
                End If
                Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, strSelectedGood)
                Parameter(1) = GenerateInputParameter("@GoodFirstCode", adVarWChar, 4000, s)
                Parameter(2) = GenerateInputParameter("@intServePlace", adVarWChar, 4000, T)
                Parameter(3) = GenerateInputParameter("@fltUsedValue", adVarWChar, 4000, U)
                Parameter(4) = GenerateInputParameter("@Pert", adVarWChar, 4000, p)

                RunParametricStoredProcedure "Insert_UsePercent", Parameter
            
                
            End With
            
    End Select
        
    MyFormAddEditMode = ViewMode 'View Mode
    SetFirstToolBar

End Sub
Public Sub Find()
    
    frmFindGoods.Show vbModal
    Dim i As Long
    i = vsGoodFirst.FindRow(mvarcode, 1, 1, True, True)
    If i > 0 Then
        vsGoodFirst.Row = i
        vsGoodFirst.ShowCell i, 0
    End If

End Sub


Public Sub SetFirstToolBar()

    Dim i As Integer
    
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(13).Enabled = True   'Find
    
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    Select Case MyFormAddEditMode
    
        Case ViewMode
        
            mdifrm.Toolbar1.Buttons(7).Enabled = True 'Edit key
            mdifrm.Toolbar1.Buttons(15).Enabled = True  'Print key
        Case AddMode
                    
            mdifrm.Toolbar1.Buttons(6).Enabled = True 'add key
            mdifrm.Toolbar1.Buttons(8).Enabled = True 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = True 'cancel key
            mdifrm.Toolbar1.Buttons(15).Enabled = False
        Case EditMode
        
            mdifrm.Toolbar1.Buttons(8).Enabled = True 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = True 'cancel key
            mdifrm.Toolbar1.Buttons(15).Enabled = False
    
    End Select
    CodeVsGood = "-1"
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
End Sub

Public Sub UpdateVsGoodFirst(Optional s As String)

    Dim i As Integer
    Dim j As Integer
    Dim Rst As New ADODB.Recordset
    
    If vsGoodFirst.Rows = 1 Then Exit Sub
    With vsGoodFirst
        .Cell(flexcpText, 1, 9, .Rows - 1, .Cols - 1) = ""
        '.Cell(flexcpText, 1, 4, .Rows - 1, 4) = 0
        For i = 1 To .Rows - 1
            .TextMatrix(i, 4) = 0
        Next
        If Trim(s) = "" Then
            Set Rst = Nothing
            Exit Sub
        End If
        
        CodeVsGoodFirst = ""
            If Rst.State <> 0 Then Rst.Close
            
            
            
            ReDim Parameter(0) As Parameter
            
            Parameter(0) = GenerateInputParameter("@nvcGoodcode", adVarWChar, 1000, s)
            
            Set Rst = RunParametricStoredProcedure2Rec("Get_UsePercent", Parameter)
                    
            If Not (Rst.EOF = True And Rst.BOF = True) Then
                While Rst.EOF = False
                    
                    For i = 1 To .Rows - 1
                        If Rst.Fields("GoodFirstcode").Value = .TextMatrix(i, 1) Then
                            For j = 9 To .Cols - 1 Step 2
                                If .Cell(flexcpData, 0, j, 0, j) = Rst.Fields("intservePlace").Value Then
                                    .TextMatrix(i, j) = Rst.Fields("fltusedvalue").Value
                                    .TextMatrix(i, j + 1) = IIf(IsNull(Rst!Pert), "", Rst!Pert)
                                    .TextMatrix(i, 4) = -1
                                    Exit For
                                End If
                            Next j
                            Exit For
                        End If
                    Next i
                    Rst.MoveNext
                Wend
            End If
        
    End With
    
    Set Rst = Nothing
    
End Sub

Private Sub cmdCalAvgBuyPrice_Click()
     If Trim(txtDateFrom.ClipText) = "" Or Trim(txtDateTo.ClipText) = "" Then
        frmDisMsg.lblMessage = " ÊÇÑíÎ ãÚÊÈÑ æÇÑÏ ßäíÏ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@NotSupportedGoodType", adInteger, 4, EnumGoodType.forSale)
    Parameter(1) = GenerateInputParameter("@nvcFromDate", adWChar, 10, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(2) = GenerateInputParameter("@nvcToDate", adWChar, 10, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))

    RunParametricStoredProcedure "Update_tGood_By_AvgBuyPrice", Parameter
    
    FillvsGood
    FWProgressBar1.Value = 0
    frmDisMsg.lblMessage = " Èå ÑæÒ ÑÓÇäí ÇäÌÇã ÔÏ "
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
    SetFirstToolBar
    txtDateFrom.Text = Mid(clsDate.shamsi(Date), 3, 2) & "/01" & "/01"
    txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    
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

    If ClsFormAccess.frmUsePercent = False Then
        Unload Me
        Exit Sub
    End If
    
    If intVersion = Min Or intVersion = Normal Then
        ShowDisMessage "ÊÚÑíÝ ÖÑíÈ ãÕÑÝ ßÇáÇåÇ ÏÑ äÓÎå åÇí íÔÑÝÊå æ ÈÇáÇÊÑ æÌæÏ ÏÇÑÏ", 1500
        Unload Me
        Exit Sub
    End If
    
    Dim Rst As New ADODB.Recordset
    
    CenterTop Me
    
    VarActForm = Me.Name
    
    MyFormAddEditMode = ViewMode 'View Mode
    SetFirstToolBar
    
    With vsGood
        
        .Rows = 1
        .Cols = 12
        .TextMatrix(0, 0) = "ÑÏíÝ"
        .TextMatrix(0, 1) = "˜Ï ˜ÇáÇ"
        .TextMatrix(0, 2) = "˜Ï ÓØÍ Çæá ˜ÇáÇ"
        .TextMatrix(0, 3) = "˜Ï ÓØÍ Ïæã ˜ÇáÇ"
        .TextMatrix(0, 4) = "ÇäÊÎÇÈ"
        .TextMatrix(0, 5) = "Ñæå ÇÕáí"
        .TextMatrix(0, 6) = "ÒíÑ Ñæå"
        .TextMatrix(0, 7) = "äÇã ˜ÇáÇ"
        .TextMatrix(0, 8) = "ÞíãÊ ÊãÇã ÔÏå"
        .TextMatrix(0, 9) = "åÒíäå ÎÊ"
        .TextMatrix(0, 10) = "åÒíäå ÐíÑÇÆí"
        .TextMatrix(0, 11) = "ÏÑÕÏ åÒíäå ÓÑÈÇÑ"
        
        .ColDataType(4) = flexDTBoolean
'        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(3) = True
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
       
       FillvsGood
    End With
    
    With vsGoodFirst
    
        .Rows = 1
        .Cols = 9
        .TextMatrix(0, 0) = "ÑÏíÝ"
        .TextMatrix(0, 1) = "˜Ï ˜ÇáÇ"
        .TextMatrix(0, 2) = "˜Ï ÓØÍ Çæá ãÇÏå Çæáíå"
        .TextMatrix(0, 3) = "˜Ï ÓØÍ Ïæã ãÇÏå Çæáíå"
        .TextMatrix(0, 4) = "ÇäÊÎÇÈ"
        .TextMatrix(0, 5) = "Ñæå ÇÕáí"
        .TextMatrix(0, 6) = "ÒíÑ Ñæå"
        .TextMatrix(0, 7) = "äÇã ãÇÏå Çæáíå"
        .TextMatrix(0, 8) = "æÇÍÏ ˜ÇáÇ"
        
        If Rst.State = 1 Then Rst.Close
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Serveplace", Parameter)
        
        If Not (Rst.EOF = True And Rst.BOF) Then
            'Rst.moveFirst
            While Rst.EOF = False
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = Rst.Fields("Description").Value
                .Cell(flexcpData, 0, .Cols - 1, 0, .Cols - 1) = Rst.Fields("intServePlace").Value
                .ColDataType(.Cols - 1) = flexDTDouble
                
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = "ÑÊ " + Rst.Fields("Description").Value
                .Cell(flexcpData, 0, .Cols - 1, 0, .Cols - 1) = Rst.Fields("intServePlace").Value
                .ColDataType(.Cols - 1) = flexDTDouble
                
                Rst.MoveNext
            Wend
        End If
        
        .ColDataType(4) = flexDTBoolean
'        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(3) = True
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        
        FillVsGoodFirst
        
    End With
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


    Set Rst = Nothing

    FillLstGoodLevel1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub
Public Sub FillLstGoodLevel1() ' it fills the lstGoodLevel1 using table tgoodlevel1
    Dim Rst As New ADODB.Recordset
    
    lstGoodLevel1.Clear
    lstGoodLevel2.Clear
    
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
           
    lstGoodLevel2.Clear
    
    
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
        
'        vsGood.ColComboList(15) = vsGood.BuildComboList(rctmp, "Description", "Code")
        
        Set Rst = Nothing
        lstGoodLevel2.ListIndex = 0
        FillVsGoodFirst
        
    End If
    
End Sub

Private Sub lstGoodLevel1_Click()

    FillLstGoodLevel2
End Sub

Private Sub lstGoodLevel1_ItemCheck(Item As Integer)
    
    Dim i As Integer
    
    If lstGoodLevel1.Selected(Item) = True Then
        For i = 0 To lstGoodLevel1.ListCount - 1
            If i <> Item And lstGoodLevel1.Selected(i) = True Then
                lstGoodLevel1.Selected(i) = False

            End If
        Next i
    End If
    FillVsGoodFirst
End Sub
Private Sub lstGoodLevel2_ItemCheck(Item As Integer)
    
    FillVsGoodFirst

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub SetUsedPercent_Click()
    
    FWProgressBar1.Value = 0
    
        RunNonParametricStoredProcedure "Update_FinalPrice"
        
         FWProgressBar1.Value = FWProgressBar1.Value + 1
        If FWProgressBar1.Value = 10 Then
           FWProgressBar1.Value = 0
        End If
        FillvsGood
        FWProgressBar1.Value = 0
        frmDisMsg.lblMessage = " Èå ÑæÒ ÑÓÇäí ÇäÌÇã ÔÏ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal

End Sub


Private Sub vsGood_AfterSort(ByVal Col As Long, Order As Integer)
    With vsGood
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next i
        End If
    End With
    
End Sub

Private Sub vsGood_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim i As Integer
    Dim s As String
    
    
    If KeyCode <> 32 Then Exit Sub
    With vsGood
        If .Rows < 2 Then Exit Sub
        If .Col = 4 And .Row > 0 Then
            .Select .Row, .Col
            .EditCell
            
            MyFormAddEditMode = ViewMode 'View Mode
            SetFirstToolBar
            
            If .TextMatrix(.Row, .Col) = -1 Then
'                For i = 1 To .Rows - 1
'                    If i <> .Row Then
'                        .TextMatrix(i, 4) = 0
'                    End If
'                Next i
                If .Row > 1 Then
                    .Cell(flexcpText, 1, 4, .Row - 1, 4) = 0
                    .Cell(flexcpText, .Row + 1, 4, .Rows - 1, 4) = 0
                Else
                    .Cell(flexcpText, .Row + 1, 4, .Rows - 1, 4) = 0
                End If
                
                s = .TextMatrix(.Row, 1)
            End If
            UpdateVsGoodFirst (s)
        End If
    End With

End Sub

Private Sub vsGood_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    Dim i As Integer
    Dim s As String
    
        With vsGood
            If .Rows < 2 Then Exit Sub
            If .Col = 4 And .Row > 0 Then
                .Select .Row, .Col
                .EditCell
                If .TextMatrix(.Row, .Col) = -1 Then
                    For i = 1 To .Rows - 1
                        If i <> .Row Then
                            .TextMatrix(i, 4) = 0
                        End If
                    Next i
'''                    If .Row > 1 Then
'''                        .Cell(flexcpText, 1, 4, .Row - 1, 4) = 0
'''                        .Cell(flexcpText, .Row + 1, 4, .Rows - 1, 4) = 0
'''                    Else
'''                        .Cell(flexcpText, .Row + 1, 4, .Rows - 1, 4) = 0
'''                    End If
                    s = .TextMatrix(.Row, 1)
                End If
                UpdateVsGoodFirst (s)
            End If
            If (.Col = 9 Or .Col = 10 Or .Col = 11) And .Row > 0 Then
                Dim Str As String
                Str = .TextMatrix(.Row, 1)
                If Str = CodeVsGood Or CodeVsGood = "0" Then
                    .Select .Row, .Col
                    .EditCell
                    CodeVsGood = .TextMatrix(.Row, 1)
                End If
            End If
        End With
        
    
End Sub


Private Sub vsGoodFirst_AfterSort(ByVal Col As Long, Order As Integer)
    With vsGoodFirst
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next i
        End If
    End With

End Sub

Private Sub vsGoodFirst_KeyDown(KeyCode As Integer, Shift As Integer)

    
    With vsGoodFirst
        Select Case MyFormAddEditMode
            Case Is <> ViewMode
            
                Dim strSelectedGood  As String
                
                With vsGood
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, 4) = -1 Then
                            strSelectedGood = .TextMatrix(i, 1)
                            Exit For
                        End If
                    Next i
                End With
                
                If strSelectedGood = "" Then
                
                    frmMsg.fwlblMsg.Caption = "áØÝÇ ÇÈÊÏÇ í˜ ˜ÇáÇ ÇäÊÎÇÈ äãÇííÏ "
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ÞÈæá"
                    frmMsg.Show vbModal
                    Exit Sub
                    
                End If
            
''                If .Col = 4 And .Row > 0 And KeyCode = 32 Then
''                    .Select .Row, .Col
''                    .EditCell
''                End If
                
                If .Col > 8 And .Row > 0 Then
                    .Select .Row, .Col
                    .EditCell
                    If .TextMatrix(.Row, 4) <> -1 Then
                        .TextMatrix(.Row, 4) = -1
                    End If
                    If IsNumeric(Chr(KeyCode)) = False And KeyCode <> 46 Then
                        KeyCode = 0
                    End If
                End If
                
        End Select
    End With

End Sub


Private Sub vsGoodFirst_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsGoodFirst
        If .Col > 8 And .Row > 0 Then
            If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 And KeyAscii <> 8 Then
                KeyAscii = 0
            End If
            If .TextMatrix(Row, 4) <> -1 Then
                .TextMatrix(Row, 4) = -1
            End If
        End If
    End With
End Sub

Private Sub vsGoodFirst_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        
        
        With vsGoodFirst
            Select Case MyFormAddEditMode
                Case Is <> ViewMode
                
                    Dim strSelectedGood  As String
                    
                    With vsGood
                        For i = 1 To .Rows - 1
                            If .TextMatrix(i, 4) = -1 Then
                                strSelectedGood = .TextMatrix(i, 1)
                                Exit For
                            End If
                        Next i
                    End With
                    
                    If strSelectedGood = "" Then
'
'                        frmMsg.FWlblMsg.Caption = "áØÝÇ ÇÈÊÏÇ í˜ ˜ÇáÇ ÇäÊÎÇÈ äãÇííÏ "
'                        frmMsg.fwBtn(0).Visible = False
'                        frmMsg.fwBtn(1).ButtonType = flwButtonOk
'                        frmMsg.fwBtn(1).Caption = "ÞÈæá"
'                        frmMsg.Show vbModal
                        Exit Sub
                        
                    End If
        
'                    If .Col = 4 And .Row > 0 Then
'                        .Select .Row, .Col
'                        .EditCell
'                    End If
                    
                    If .Col > 8 And .Row > 0 Then
                        If .TextMatrix(.Row, 4) <> -1 Then
                            .TextMatrix(.Row, 4) = -1
                        End If
                    
                        .Select .Row, .Col
                        .EditCell
                    End If
            End Select
        End With

End Sub

Private Sub vsGoodFirst_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With vsGoodFirst
        .Row = Row
        .Col = Col
    End With

End Sub
Public Sub Printing()
    On Error GoTo Err_Handler
             
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
        
        
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepGoodsWithFirstGoods.rpt"
        CrystalReport1.ReportTitle = "ÒÇÑÔ ßÇáÇåÇ æ ãæÇÏ ãÕÑÝí"
        Dim fileSystem As New FileSystemObject
        Dim IsFileExist As Boolean
        IsFileExist = fileSystem.FileExists(CrystalReport1.ReportFileName)
        If IsFileExist = False Then
                frmDisMsg.lblMessage = " ÝÇíá  " & CrystalReport1.ReportFileName & "íÏÇ äÔÏ "
                frmDisMsg.Timer1.Interval = 3000
                frmDisMsg.Timer1.Enabled = True
                frmDisMsg.Show vbModal
                Exit Sub
        End If
        
     
        Dim intIndex As Integer
        
        For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
            CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
        Next intIndex
      
        CrystalReport1.Destination = crptToWindow 'crptToPrinter '
        CrystalReport1.RetrieveDataFiles
        ODBCSetting clsArya.ServerName, clsArya.DbName
        CrystalReport1.Connect = CrystallConnection
        CrystalReport1.Action = 1
        
        If Screen.Width > 12000 Then
            CrystalReport1.PageZoom (100)
        Else
            CrystalReport1.PageZoom (75)
        End If
 

Exit Sub
Err_Handler:
    LogSaveNew "frmUsePercent => ", err.Description, err.Number, err.Source, "Printing"
    ShowErrorMessage
End Sub
