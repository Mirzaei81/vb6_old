VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmFindSendDeliveries 
   BackColor       =   &H00E0E0E0&
   Caption         =   $"frmFindSendDeliveries.frx":0000
   ClientHeight    =   9510
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   13500
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   12
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H000000C0&
   Icon            =   "frmFindSendDeliveries.frx":00C8
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9510
   ScaleWidth      =   13500
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   720
      Top             =   0
   End
   Begin VB.TextBox txtTel 
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
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   360
      Width           =   2355
   End
   Begin VB.TextBox txtName 
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
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   360
      Width           =   2355
   End
   Begin VB.Frame Frame_Customers 
      Caption         =   "Ã” ÃÊÌ „‘ —Ì«‰"
      Height          =   5055
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1920
      Width           =   10335
      Begin VB.CommandButton cmdClose 
         Caption         =   "»” ‰"
         Height          =   435
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   4560
         Width           =   1215
      End
      Begin VSFlex7LCtl.VSFlexGrid vsCustomers 
         Height          =   4005
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   10035
         _cx             =   17701
         _cy             =   7064
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   16761024
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   32896
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   500
         RowHeightMax    =   500
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFindSendDeliveries.frx":A58A
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
         FillStyle       =   1
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
   End
   Begin VB.ComboBox CmbPayk 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7200
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "›Ì· — ﬂ—œ‰ ›«ﬂ Ê—Â«Ì «—”«·Ì  Ê”ÿ ÅÌﬂ Â«"
      Top             =   960
      Width           =   2850
   End
   Begin VB.TextBox txtNo 
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
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1395
   End
   Begin VB.CheckBox ChkDaily 
      Alignment       =   1  'Right Justify
      Caption         =   "«—”«·Ì Â«Ì «„—Ê“"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   960
      Value           =   1  'Checked
      Width           =   2145
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
      Height          =   615
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   8760
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
      TabIndex        =   3
      Top             =   8760
      Width           =   1215
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmFindSendDeliveries.frx":A64E
      TabIndex        =   10
      Top             =   0
      Width           =   480
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFactors 
      Height          =   7125
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   13275
      _cx             =   23416
      _cy             =   12568
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
      BackColor       =   -2147483644
      ForeColor       =   12582912
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16744576
      BackColorAlternate=   -2147483644
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFindSendDeliveries.frx":A6D4
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
   Begin FLWCtrls.FWNumericTextBox txtInterval 
      Height          =   480
      Left            =   720
      TabIndex        =   21
      Top             =   960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   847
      Min             =   1
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   3225
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   510
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "«Ì‰ —‰ "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2445
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   495
      Width           =   735
   End
   Begin VB.Label lblSecondTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "À«‰ÌÂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1005
      Width           =   495
   End
   Begin VB.Label lblIntervalTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "“„«‰ »—Ê“ —”«‰Ì :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1005
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰"
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
      Height          =   375
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„"
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
      Height          =   375
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ÅÌﬂ Â« "
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
      Height          =   375
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "»Ì—Ê‰"
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
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblPayk 
      Alignment       =   2  'Center
      Caption         =   "œ·ÌÊ—Ì"
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
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblColorOut 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblColorPayk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "«‘ —«ò"
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
      Height          =   375
      Left            =   12480
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Label LblFindFactor 
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
      Height          =   615
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   5055
   End
End
Attribute VB_Name = "frmFindSendDeliveries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Long
Dim FactorType As EnumFactorType
Dim Parameter() As Parameter

Private Sub CancelButton_Click()
    mvarcode = 0
    Unload Me
End Sub

Private Sub ChkDaily_Click()
    FillvsFactors
End Sub


Private Sub CmbPayk_Click()
    If CmbPayk.ListIndex <> -1 Then FillvsFactors
End Sub

Private Sub cmdClose_Click()
    Frame_Customers.Visible = False
End Sub

Private Sub Form_Activate()
    txtTel.SetFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    ElseIf KeyCode = 13 Then
        OKButton_Click
    ElseIf KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Or KeyCode = 41 Then
        vsFactors.SetFocus
    End If

End Sub
Private Sub FillsPaykCombo()
    Dim rctmp As New ADODB.Recordset
    CmbPayk.Clear
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Job", adInteger, 4, EnumIncharge.Payk)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_vw_Per_BY_Job", Parameter)
    CmbPayk.AddItem "Â„Â ÅÌﬂ Â«"
    CmbPayk.ItemData(0) = 0
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
        
            CmbPayk.AddItem CStr(rctmp.Fields("nvcFirstName")) & " " & CStr(rctmp.Fields("nvcSurName"))
            CmbPayk.ItemData(CmbPayk.ListCount - 1) = Val(rctmp.Fields("pPNo"))
            rctmp.MoveNext
        Loop
    End If
    CmbPayk.ListIndex = 0
    rctmp.Close
    Set rctmp = Nothing
End Sub

Private Sub Form_Load()

    CenterCenterinSecondScreen Me
    
    FactorType = EnumFactorType.Invoice
    mvarcode = 0
   With vsFactors
        .Rows = 1
        .Cols = 17
        .ColAlignment(-1) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColWidth(1) = 500
        '.ColDataType(1) = flexDTBoolean
        .ColHidden(0) = True
        .ColHidden(1) = False
        .TextMatrix(0, 0) = "òœ ›Ì‘"
        .TextMatrix(0, 1) = "ÅÌﬂ"
        .TextMatrix(0, 2) = "”—Ì«·"
        .TextMatrix(0, 3) = "«Œÿ«—"
        .TextMatrix(0, 4) = "«‘ —«ò"
        .TextMatrix(0, 5) = "„‘ —Ì"
        .TextMatrix(0, 6) = "„»·€"
        .TextMatrix(0, 7) = "”«⁄  ”›«—‘"
        .TextMatrix(0, 8) = "„œ  ”›«—‘"
        .TextMatrix(0, 9) = "”«⁄  «—”«·"
        .TextMatrix(0, 10) = "„œ  «—”«·"
        .TextMatrix(0, 11) = " «—ÌŒ"
        .TextMatrix(0, 12) = "‰Ê⁄"
        .TextMatrix(0, 13) = "Ê÷⁄Ì "
        .TextMatrix(0, 14) = "¬œ—”"
        .TextMatrix(0, 15) = "¬œ—” „Êﬁ "
        .TextMatrix(0, 16) = " ·›‰ "
        
        .AutoSearch = flexSearchFromCursor
    
    End With
   With vsCustomers
        .Rows = 1
        .Cols = 9
        .ColAlignment(-1) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColHidden(0) = True
        .ColHidden(1) = False
        .ColWidth(1) = 500
        '.ColDataType(1) = flexDTBoolean
        .TextMatrix(0, 1) = "ÅÌﬂ"
        .TextMatrix(0, 2) = "„‘ —Ì"
        .TextMatrix(0, 3) = " ·›‰"
        .TextMatrix(0, 4) = "”«⁄  ”›«—‘"
        .TextMatrix(0, 5) = "„œ  ”›«—‘"
        .TextMatrix(0, 6) = "”«⁄  «—”«·"
        .TextMatrix(0, 7) = "„œ  «—”«·"
        .TextMatrix(0, 8) = "Ê÷⁄Ì "
        
        .AutoSearch = flexSearchFromCursor
    
    End With
    If Val(GetSetting(strMainKey, Me.Name, "TimerInterval")) > 0 Then
        Timer1.Interval = Val(GetSetting(strMainKey, Me.Name, "TimerInterval"))
    Else
         Timer1.Interval = 10000
    End If
    
    txtInterval.Value = CStr(Timer1.Interval / 1000)
    FillsPaykCombo
    FillvsFactors
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
    
    Frame_Customers.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    SaveSetting strMainKey, Me.Name, "TimerInterval", CStr(Val(txtInterval.Value) * 1000)
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top


    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub


Private Sub OKButton_Click()
    If vsFactors.Row > 0 Then
        mvarcode = vsFactors.TextMatrix(vsFactors.Row, 2)
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
Private Sub FillCustomersByName()
    Timer1.Enabled = False
    Dim i As Long
    Frame_Customers.Visible = True
    With vsCustomers
    .Rows = 1
    For i = 1 To vsFactors.Rows - 1
        If InStr(1, vsFactors.TextMatrix(i, 5), Trim(TxtName), vbTextCompare) > 0 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = vsFactors.TextMatrix(i, 1)
            .TextMatrix(.Rows - 1, 2) = vsFactors.TextMatrix(i, 5)
            .TextMatrix(.Rows - 1, 3) = vsFactors.TextMatrix(i, 16)
            .TextMatrix(.Rows - 1, 4) = vsFactors.TextMatrix(i, 7)
            .TextMatrix(.Rows - 1, 5) = vsFactors.TextMatrix(i, 8)
            .TextMatrix(.Rows - 1, 6) = vsFactors.TextMatrix(i, 9)
            .TextMatrix(.Rows - 1, 7) = vsFactors.TextMatrix(i, 10)
            .TextMatrix(.Rows - 1, 8) = vsFactors.TextMatrix(i, 13)
        
        End If
    Next
    End With
    Timer1.Enabled = True

End Sub
Private Sub FillCustomersByTel()
    Timer1.Enabled = False
    Dim i As Long
    Frame_Customers.Visible = True
    With vsCustomers
    .Rows = 1
    For i = 1 To vsFactors.Rows - 1
        If InStr(1, vsFactors.TextMatrix(i, 16), Trim(txtTel), vbTextCompare) > 0 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = vsFactors.TextMatrix(i, 1)
            .TextMatrix(.Rows - 1, 2) = vsFactors.TextMatrix(i, 5)
            .TextMatrix(.Rows - 1, 3) = vsFactors.TextMatrix(i, 16)
            .TextMatrix(.Rows - 1, 4) = vsFactors.TextMatrix(i, 7)
            .TextMatrix(.Rows - 1, 5) = vsFactors.TextMatrix(i, 8)
            .TextMatrix(.Rows - 1, 6) = vsFactors.TextMatrix(i, 9)
            .TextMatrix(.Rows - 1, 7) = vsFactors.TextMatrix(i, 10)
            .TextMatrix(.Rows - 1, 8) = vsFactors.TextMatrix(i, 13)
        
        End If
    Next
    End With
    Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = txtInterval.Value * 1000
    FillvsFactors

End Sub

Private Sub txtName_Change()
    
    i = -1
    vsCustomers.Rows = 1
    If Len(TxtName.Text) <> 0 Then
        i = vsFactors.FindRow(Trim(TxtName.Text), 1, 5, False, False)
        FillCustomersByName
    End If
    
    If i > 0 Then
        vsFactors.Row = i
        vsFactors.ShowCell i, 0
        LblFindFactor.Caption = ""
        vsFactors.TopRow = i
    Else
        vsFactors.Row = 0
        vsFactors.ShowCell 0, 0
        If Val(TxtName.Text) > 0 Then
           LblFindFactor.Caption = "»—«Ì ‰«„ - " & Val(TxtName.Text) & "  - ›«ò Ê— «—”«· ‘œÂ ‰œ«—Ì„"
         Else
            LblFindFactor.Caption = "‰«„ „‘ —Ì —« Ê«—œ ﬂ‰Ìœ  "
         End If
    
    End If

End Sub

Private Sub txtName_GotFocus()
    vsFactors.Row = 0
    vsFactors.Select vsFactors.Row, 5
    vsFactors.Sort = flexSortGenericAscending
  '  vsFactors.Sort = flexSortGenericDescending
    LblFindFactor.Caption = "‰«„ „‘ —Ì —« Ê«—œ ﬂ‰Ìœ  "
    For i = 1 To vsFactors.Rows - 1
        vsFactors.TextMatrix(i, 0) = i
    Next

End Sub

Private Sub txtNo_Change()
    
    i = -1
    If Len(txtNo.Text) <> 0 Then
       i = vsFactors.FindRow(txtNo.Text, 1, 4, True, True)
    End If
    If i > 0 Then
        vsFactors.Row = i
        vsFactors.ShowCell i, 0
        LblFindFactor.Caption = ""
        vsFactors.TopRow = i
    Else
        vsFactors.Row = 0
        vsFactors.ShowCell 0, 0
        If Val(txtNo.Text) > 0 Then
           LblFindFactor.Caption = "»—«Ì «‘ —«ò  " & Val(txtNo.Text) & " ›«ò Ê— «—”«· ‘œÂ ‰œ«—Ì„"
         Else
            LblFindFactor.Caption = "‘„«—Â «‘ —«ò —« Ê«—œ ﬂ‰Ìœ  "
         End If
    
    End If
    
End Sub

Private Sub txtNo_GotFocus()

    vsFactors.Row = 0
    vsFactors.Select vsFactors.Row, 4
    vsFactors.Sort = flexSortGenericAscending
  '  vsFactors.Sort = flexSortGenericDescending
    LblFindFactor.Caption = "‘„«—Â «‘ —«ò —« Ê«—œ ﬂ‰Ìœ  "
    For i = 1 To vsFactors.Rows - 1
        vsFactors.TextMatrix(i, 0) = i
    Next
    
End Sub


Private Sub FillvsFactors()

On Error GoTo ErrHandler

    Dim Rst As New ADODB.Recordset
    
    If Rst.State = 1 Then Rst.Close
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(1) = GenerateInputParameter("@Job", adInteger, 4, EnumIncharge.Payk)
    Set Rst = RunParametricStoredProcedure2Rec("GetTotal_Delivers", Parameter)
    
    Dim intServePlace As Integer
    Dim intDistance As Integer
    Dim intWarn As Integer
    With vsFactors
        .Rows = 1
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            'Rst.MoveFirst
            Dim VarToday As String
            VarToday = mvarDate
            i = 1
            While Rst.EOF = False
                If ChkDaily.Value = False Or Rst.Fields("Date").Value = VarToday Then
                    If Rst.Fields("Incharge").Value = CmbPayk.ItemData(CmbPayk.ListIndex) Or CmbPayk.ItemData(CmbPayk.ListIndex) = 0 Then
                        .Rows = .Rows + 1
                        .TextMatrix(i, 0) = Rst.Fields("intSerialNo").Value
                        .TextMatrix(i, 1) = Rst.Fields("InchargeName").Value
                        .TextMatrix(i, 2) = Rst.Fields("NO").Value ' Val(Right(Rst.Fields("No").Value, 3))
                        .TextMatrix(i, 3) = ""
                        Dim intCode As Long
                        intCode = IIf(IsNull(Rst.Fields("Code").Value), "-1", Rst.Fields("Code").Value)
                        If intCode = -1 Then
                            .TextMatrix(i, 4) = ""
                        Else
                            .TextMatrix(i, 4) = intCode
                        End If
                        .TextMatrix(i, 5) = IIf(IsNull(Rst.Fields("Full Name").Value), " ", Rst.Fields("Full Name").Value)
                        .TextMatrix(i, 6) = Rst.Fields("SumPrice").Value
                        .TextMatrix(i, 7) = Rst.Fields("Time").Value
                        .TextMatrix(i, 8) = Rst.Fields("RemainTime").Value
                        .TextMatrix(i, 9) = Rst.Fields("TimeSend").Value
                        .TextMatrix(i, 10) = Rst.Fields("RemainTimeSend").Value
                        .TextMatrix(i, 11) = Rst.Fields("Date").Value
                        .TextMatrix(i, 12) = Rst.Fields("ServePlaceName").Value
                        .TextMatrix(i, 13) = Rst.Fields("DeliverStatus").Value
                        .TextMatrix(i, 14) = IIf(IsNull(Rst.Fields("Address").Value), " ", Rst.Fields("Address").Value)
                        .TextMatrix(i, 15) = IIf(IsNull(Rst.Fields("TempAddress").Value), "", Rst.Fields("TempAddress").Value)
                        .TextMatrix(i, 16) = IIf(IsNull(Rst.Fields("TelNumber").Value), "", Rst.Fields("TelNumber").Value)
                        intServePlace = Rst.Fields("ServePlace").Value
                        intDistance = IIf(IsNull(Rst.Fields("distance").Value), 0, Rst.Fields("distance").Value)
                        If intServePlace = 2 Then
                            If Rst.Fields("StationId").Value = -1 Then
                                .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &H4080&
                            Else
                                .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HC0C0&
                            End If
                        ElseIf intServePlace = 4 Then
                            .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HFF80FF
                        End If
                         intWarn = Rst.Fields("intWarn").Value
                        If intWarn = -1 Then
                            .Cell(flexcpBackColor, .Rows - 1, 3, .Rows - 1, 3) = &HFF&
                        End If
                        i = i + 1
                    End If
                End If
                
                Rst.MoveNext
            Wend
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing

Exit Sub
ErrHandler:

     ShowDisMessage "StationId ____" & err.Description, 1000

End Sub


Private Sub txtTel_Change()
    i = -1
    vsCustomers.Rows = 1
    If Len(txtTel.Text) <> 0 Then
        i = vsFactors.FindRow(Trim(txtTel.Text), 1, 16, False, False)
        FillCustomersByTel
    End If
    
    If i > 0 Then
        vsFactors.Row = i
        vsFactors.ShowCell i, 0
        LblFindFactor.Caption = ""
        vsFactors.TopRow = i
    Else
        vsFactors.Row = 0
        vsFactors.ShowCell 0, 0
        If Val(TxtName.Text) > 0 Then
           LblFindFactor.Caption = "»—«Ì  ·›‰ - " & Val(txtTel.Text) & "  - ›«ò Ê—  ‰œ«—Ì„"
         Else
            LblFindFactor.Caption = " ·›‰ „‘ —Ì —« Ê«—œ ﬂ‰Ìœ  "
         End If
    End If

End Sub

Private Sub txtTel_GotFocus()
    vsFactors.Row = 0
    vsFactors.Select vsFactors.Row, 16
    vsFactors.Sort = flexSortGenericAscending
  '  vsFactors.Sort = flexSortGenericDescending
    LblFindFactor.Caption = " ·›‰ „‘ —Ì —« Ê«—œ ﬂ‰Ìœ  "
    For i = 1 To vsFactors.Rows - 1
        vsFactors.TextMatrix(i, 0) = i
    Next

End Sub

Private Sub vsFactors_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsFactors.Rows - 1
        vsFactors.TextMatrix(i, 0) = i
    Next
    
End Sub

Private Sub vsFactors_DblClick()
    If vsFactors.Row > 0 Then
        OKButton_Click
    End If
End Sub

Private Sub Define_Factor()

    
''''    Dim Rst As New ADODB.Recordset
''''
''''    ReDim Parameter(3) As Parameter
''''    Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, FactorType)
''''    Parameter(1) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
''''    Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
''''    Parameter(3) = GenerateInputParameter("@No", adBigInt, 8, txtNo.Text)
''''
''''    Set Rst = RunParametricStoredProcedure2Rec("Get_Define_Factors", Parameter)
''''
''''    With vsFactors
''''        .Rows = 1
''''        .Cols = 8
''''        i = 0
''''        While Rst.EOF <> True
''''            .Rows = .Rows + 1
''''            i = .Rows - 1
''''            .TextMatrix(i, 0) = i
''''            .TextMatrix(i, 1) = Rst!intSerialNo
''''            .TextMatrix(i, 2) = Rst!No
''''            .TextMatrix(i, 3) = Rst![Name]
''''            .TextMatrix(i, 4) = Rst!Amount
''''            .TextMatrix(i, 5) = Rst!FeeUnit * Rst!Amount
''''            .TextMatrix(i, 6) = Rst!nvcFirstName & " " & Rst!nvcSurName
''''   '         .TextMatrix(i, 7) = Right(Rst!No, 3)
''''            Rst.MoveNext
''''
''''        Wend
''''    End With
    
''''    Set Rst = Nothing


End Sub

