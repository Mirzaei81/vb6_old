VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmShowInvoiceMenu 
   Caption         =   "                     "
   ClientHeight    =   6585
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   9975
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Nazanin"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShowInvoiceMenu.frx":0000
   KeyPreview      =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   9975
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4200
      Width           =   9735
      Begin VB.Label lblDiscountTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label LblDiscount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Œ›Ì›"
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
         Height          =   375
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LblCarryFee 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ò—«ÌÂ Õ„·"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblCarryFeeTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label LblSubTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ã„⁄     "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label LblDutyTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDuty 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄Ê«—÷"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LblService 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”—ÊÌ”"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblServiceTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblPackingTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label LblPacking 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»” Â »‰œÌ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label LblTax 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„«·Ì« "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblTaxTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblSumPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   795
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ã„⁄  ò·   "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmShowInvoiceMenu.frx":B16C
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
   Begin VSFlex7LCtl.VSFlexGrid vsInvoice 
      Height          =   4035
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   9795
      _cx             =   17277
      _cy             =   7117
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
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   8
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÿ—Õ Ê «Ã—« : ê—ÊÂ ‘—ò Â«Ì ¬—Ì«             www.FGArya.Com  "
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   5880
      Width           =   6375
   End
   Begin VB.Image Image_Fgarya 
      Height          =   1035
      Left            =   240
      Picture         =   "frmShowInvoiceMenu.frx":B1F2
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   2775
   End
End
Attribute VB_Name = "frmShowInvoiceMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Result As Boolean
Dim i As Integer
Private rctmp As New ADODB.Recordset
Dim No As Double
Dim intSerialNo As Double

Private Sub Form_Activate()
    Me.Caption = Space(50) & "›«ò Ê—›—Ê‘"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' Select Case Shift
'      Case 0
'          Select Case KeyCode
'            Case 13
'              SendKeys "{Tab}", 12
'                Case 113  ' F2
'
'                 End Select
'    End Select
End Sub

Private Sub Form_Load()

    CenterCenter Me
    
    Result = False
      
'    GetDataDetail
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
    With vsInvoice
        .Cols = 12
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "„ﬁœ«—"
        .TextMatrix(0, 2) = "‰«„ ò«·«"
        .TextMatrix(0, 3) = "›Ì"
        .TextMatrix(0, 4) = "Ã„⁄"
        .TextMatrix(0, 5) = "ﬂœ ﬂ«·«"
        .TextMatrix(0, 6) = "”—Ê"
        .TextMatrix(0, 7) = " €ÌÌ—« "
        .TextMatrix(0, 8) = "œ—’œ  Œ›Ì›"
        .TextMatrix(0, 9) = "‰—Œ"
        .TextMatrix(0, 10) = "⁄Ê«—÷"
        .TextMatrix(0, 11) = "„«·Ì« "
        
        .ColAlignment(-1) = flexAlignCenterCenter
       ' .ColAlignment(2) = flexAlignRightCenter
        
        .ColDataType(10) = flexDTBoolean
        .ColDataType(11) = flexDTBoolean
    
         For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, "frmShowInvoice_vsInvoice", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = .Width / 10     'Row
            End If
         Next i
        Dim Rst As New ADODB.Recordset
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Serveplace", Parameter)
        
        .ColComboList(6) = .BuildComboList(Rst, "Description", "intServePlace")
        If Rst.State <> 0 Then Rst.Close
    
    End With
    On Error GoTo ErrHandler
'    Dim LogoFile As String
'    Dim filetemp As New FileSystemObject
'    LogoFile = App.Path & "\Image\Logo.gif"
'    If filetemp.FileExists(LogoFile) Then
'        Image1.Picture = LoadPicture(LogoFile)
'    Else
'        LogoFile = App.Path & "\Image\Logo.jpg"
'        If filetemp.FileExists(LogoFile) Then
'            Image1.Picture = LoadPicture(LogoFile)
'        End If
'    End If
    
Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 1500
    Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

'    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
        vsInvoice.RowHeightMax = vsInvoice.Height / (8 * 1.08) '8.2
        vsInvoice.RowHeightMin = vsInvoice.Height / (8 * 1.11)  '8.5
    End If

End Sub

Public Sub UpdateLblValue()
    
    lblCarryFeeTotal.Caption = frmInvoice.lblCarryFeeTotal
    lblDiscountTotal.Caption = frmInvoice.lblDiscountTotal
    lblPackingTotal.Caption = frmInvoice.lblPackingTotal
    lblServiceTotal.Caption = frmInvoice.lblServiceTotal
    LblSubTotal.Caption = frmInvoice.LblSubTotal
    lblTaxTotal.Caption = frmInvoice.lblTaxTotal
    LblDutyTotal.Caption = frmInvoice.LblDutyTotal
    
    lblSumPrice.Caption = frmInvoice.lblSumPrice.Caption
End Sub
Public Sub UpdateGridValue()
    Dim i As Long
    With vsInvoice
        .Rows = frmInvoice.FlxDetail.Rows
        For i = 1 To frmInvoice.FlxDetail.Rows - 1
            
            .TextMatrix(i, 0) = frmInvoice.FlxDetail.TextMatrix(i, 0)
            .TextMatrix(i, 1) = frmInvoice.FlxDetail.TextMatrix(i, 1)
            .TextMatrix(i, 2) = frmInvoice.FlxDetail.TextMatrix(i, 2)
            .TextMatrix(i, 3) = frmInvoice.FlxDetail.TextMatrix(i, 3)
            .TextMatrix(i, 4) = frmInvoice.FlxDetail.TextMatrix(i, 4)
            .TextMatrix(i, 5) = frmInvoice.FlxDetail.TextMatrix(i, 5)
            .TextMatrix(i, 6) = frmInvoice.FlxDetail.TextMatrix(i, 8)
            .TextMatrix(i, 7) = Trim(frmInvoice.FlxDetail.TextMatrix(i, 10))
            .TextMatrix(i, 8) = frmInvoice.FlxDetail.TextMatrix(i, 11)
            .TextMatrix(i, 9) = frmInvoice.FlxDetail.TextMatrix(i, 12)
            .TextMatrix(i, 10) = frmInvoice.FlxDetail.TextMatrix(i, 17)
            .TextMatrix(i, 11) = frmInvoice.FlxDetail.TextMatrix(i, 18)
        Next
    End With
End Sub

Public Sub ClearGridValue()
    
    With vsInvoice
        .Rows = 1
        .Rows = MaxInvoiceRows
    End With
    lblCarryFeeTotal.Caption = ""
    lblDiscountTotal.Caption = ""
    lblPackingTotal.Caption = ""
    lblServiceTotal.Caption = ""
    LblSubTotal.Caption = ""
    lblTaxTotal.Caption = ""
    LblDutyTotal.Caption = ""
    
    lblSumPrice.Caption = ""

End Sub

Private Sub vsInvoice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    With vsInvoice
        If Col >= 0 Then
            For i = 0 To .Cols - 1
                SaveSetting strMainKey, "frmShowInvoice_vsInvoice", "Col" & i, .ColWidth(i)
            Next
        End If
    End With

End Sub

