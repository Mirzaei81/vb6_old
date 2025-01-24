VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmDistance 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   Icon            =   "frmDistance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   7065
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2040
      Width           =   6855
      Begin VB.CommandButton cmd_UpdatePaykFeePercent 
         BackColor       =   &H0000C0C0&
         Caption         =   " €ÌÌ— ﬂ—«ÌÂ ÅÌﬂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmd_UpdateCarryFeePercent 
         BackColor       =   &H0000C0C0&
         Caption         =   " €ÌÌ— ﬂ—«ÌÂ Õ„·"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   120
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtCarryFeePercent 
         Height          =   525
         Left            =   3720
         TabIndex        =   28
         ToolTipText     =   " ⁄œ«œ ‰›—« Ì ﬂÂ «“ „Ì“ «” ›«œÂ „Ìﬂ‰‰œ"
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   926
         _Version        =   393216
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPaykFeePercent 
         Height          =   525
         Left            =   3720
         TabIndex        =   29
         ToolTipText     =   " ⁄œ«œ ‰›—« Ì ﬂÂ «“ „Ì“ «” ›«œÂ „Ìﬂ‰‰œ"
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   926
         _Version        =   393216
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "œ—’œ  €ÌÌ—"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   390
         Index           =   2
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "œ—’œ  €ÌÌ—"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   390
         Index           =   1
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ—«ÌÂ ÅÌﬂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ—«ÌÂ Õ„·"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   3240
      Width           =   6855
      Begin VB.TextBox txtDescription 
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
         Height          =   450
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox txtCarryFee 
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
         Height          =   450
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtPaykFee 
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
         Height          =   450
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdUpdateDistance 
         BackColor       =   &H0000C0C0&
         Caption         =   "«Œ ’«’ ﬂ—«ÌÂ »Â „‘ —ﬂÌ‰ „ÕœÊœÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblDescription 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕœÊœÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblCarryFee 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ—«ÌÂ Õ„·"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   390
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label lblPaykFee 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ—«ÌÂ ÅÌﬂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   390
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   6855
      Begin VB.TextBox txtOldCarryFee 
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
         Height          =   450
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtNewCarryFee 
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
         Height          =   450
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtOldPaykFee 
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
         Height          =   450
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtNewPaykFee 
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
         Height          =   450
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdateCarryFee 
         BackColor       =   &H0000C0C0&
         Caption         =   " €ÌÌ— ﬂ—«ÌÂ Õ„·"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdatePaykFee 
         BackColor       =   &H0000C0C0&
         Caption         =   " €ÌÌ— ﬂ—«ÌÂ ÅÌﬂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ—«ÌÂ Õ„·"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   " »œÌ· ‘Êœ »Â"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   390
         Index           =   0
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ—«ÌÂ ÅÌﬂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   " »œÌ· ‘Êœ »Â"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   390
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsDistance 
      Height          =   3555
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   6915
      _cx             =   12197
      _cy             =   6271
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
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   5520
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄—Ì› „ÕœÊœÂ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmDistance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter

Public Sub SetFirstToolBar()
    Dim i As Integer

    AllButton vbOff, True
   
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    txtDescription.Locked = True
    txtCarryFee.Locked = True
    txtPaykFee.Locked = True
    cmdUpdateDistance.Enabled = True
    
    If MyFormAddEditMode = ViewMode Then  ' View Mode
 
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = True 'Delete
        
        txtDescription.Locked = True
        txtCarryFee.Locked = True
        txtPaykFee.Locked = True
        cmdUpdateDistance.Enabled = True
                
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
                
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False 'Delete
        
        txtDescription.Locked = False
        txtCarryFee.Locked = False
        txtPaykFee.Locked = False
        cmdUpdateDistance.Enabled = False
        
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False 'Delete
        
        txtDescription.Locked = False
        txtCarryFee.Locked = False
        txtPaykFee.Locked = False
        cmdUpdateDistance.Enabled = False
        
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
End Sub

Private Sub cmd_UpdateCarryFeePercent_Click()
    If Val(txtCarryFeePercent.Text) = 0 Then
        frmMsg.fwlblMsg.Caption = "„ﬁœ«— œ—’œ ﬂ—«ÌÂ Õ„·  Œ«·Ì «”  ·ÿ›« ¬‰ —« Ê«—œ ﬂ‰Ìœ"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
     
    If Val(txtCarryFeePercent.Text) < -100 Then
        frmMsg.fwlblMsg.Caption = "„ﬁœ«— œ—’œ ﬂ—«ÌÂ Õ„·   ‰»«Ìœ ò„ — «“ 100 œ—’œ „‰›Ì »«‘œ"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    frmMsg.fwlblMsg.Caption = " ¬Ì« „Ì ŒÊ«ÂÌœ ﬂ—«ÌÂ Õ„· ÃœÌœ Ã«Ìê“Ì‰ ﬂ—«ÌÂ Õ„· ﬁ»·Ì ê—œœ "
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(0).Caption = "»·Ì"
    frmMsg.fwBtn(1).ButtonType = flwButtonNo
    frmMsg.fwBtn(1).Caption = "ŒÌ—"
    frmMsg.Show vbModal
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
 
     
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@OldCarryFee", adDouble, 8, 0)
    Parameter(1) = GenerateInputParameter("@NewCarryFee", adDouble, 8, 0)
    Parameter(2) = GenerateInputParameter("@PercentCarryFee", adDouble, 8, Val(txtCarryFeePercent.Text))
    Parameter(3) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
    Parameter(4) = GenerateOutputParameter("@Updated", adBigInt, 8)
                          
    Dim Updated As Long
    Updated = RunParametricStoredProcedure("Update_Cust_By_NewCarryFee", Parameter)
    If Updated > 0 Then ShowDisMessage " €ÌÌ—«  «‰Ã«„ ‘œ", 1000 Else ShowDisMessage "œ— À»   €ÌÌ—«  „‘ò· ÊÃÊœ œ«—œ", 2000
    DefaultSetting

End Sub

Private Sub cmd_UpdatePaykFeePercent_Click()
    If Val(txtPaykFeePercent.Text) = 0 Then
        frmMsg.fwlblMsg.Caption = "„ﬁœ«— œ—’œ ﬂ—«ÌÂ ÅÌﬂ  Œ«·Ì «”  ·ÿ›« ¬‰ —« Ê«—œ ﬂ‰Ìœ"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
     
    If Val(txtPaykFeePercent.Text) < -100 Then
        frmMsg.fwlblMsg.Caption = "„ﬁœ«— œ—’œ ﬂ—«ÌÂ ÅÌﬂ  ‰»«Ìœ ò„ — «“ 100 œ—’œ „‰›Ì »«‘œ"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    frmMsg.fwlblMsg.Caption = " ¬Ì« „Ì ŒÊ«ÂÌœ ﬂ—«ÌÂ ÅÌﬂ ÃœÌœ Ã«Ìê“Ì‰ ﬂ—«ÌÂ ÅÌﬂ ﬁ»·Ì ê—œœ "
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(0).Caption = "»·Ì"
    frmMsg.fwBtn(1).ButtonType = flwButtonNo
    frmMsg.fwBtn(1).Caption = "ŒÌ—"
    frmMsg.Show vbModal
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
 
     
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@OldPaykFee", adDouble, 8, 0)
    Parameter(1) = GenerateInputParameter("@NewPaykFee", adDouble, 8, 0)
    Parameter(2) = GenerateInputParameter("@PercentPaykFee", adDouble, 8, Val(txtPaykFeePercent.Text))
    Parameter(3) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
    Parameter(4) = GenerateOutputParameter("@Updated", adBigInt, 8)
                          
    Dim Updated As Long
    Updated = RunParametricStoredProcedure("Update_Cust_By_NewPaykFee", Parameter)
    If Updated > 0 Then ShowDisMessage " €ÌÌ—«  «‰Ã«„ ‘œ", 1000 Else ShowDisMessage "œ— À»   €ÌÌ—«  „‘ò· ÊÃÊœ œ«—œ", 2000
    DefaultSetting

End Sub

Private Sub cmdUpdateCarryFee_Click()
    If Val(txtNewCarryFee.Text) < 0 Or Val(txtOldCarryFee.Text) < 0 Then
        frmMsg.fwlblMsg.Caption = "„ﬁœ«— ﬂ—«ÌÂ Õ„· ‰„Ì  Ê«‰œ ò„ — «“ ’›— »«‘œ "
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    If txtNewCarryFee.Text = "" Or txtOldCarryFee.Text = "" Then
        frmMsg.fwlblMsg.Caption = "„ﬁœ«— ﬂ—«ÌÂ Õ„·  Œ«·Ì «”  ·ÿ›« ﬂ—«ÌÂ Õ„· —« Ê«—œ ﬂ‰Ìœ"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
     
    frmMsg.fwlblMsg.Caption = " ¬Ì« „Ì ŒÊ«ÂÌœ ﬂ—«ÌÂ Õ„· ÃœÌœ Ã«Ìê“Ì‰ ﬂ—«ÌÂ Õ„· ﬁ»·Ì ê—œœ "
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(0).Caption = "»·Ì"
    frmMsg.fwBtn(1).ButtonType = flwButtonNo
    frmMsg.fwBtn(1).Caption = "ŒÌ—"
    frmMsg.Show vbModal
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
 
     
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@OldCarryFee", adDouble, 8, Val(txtOldCarryFee.Text))
    Parameter(1) = GenerateInputParameter("@NewCarryFee", adDouble, 8, Val(txtNewCarryFee.Text))
    Parameter(2) = GenerateInputParameter("@PercentCarryFee", adDouble, 8, 0)
    Parameter(3) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
    Parameter(4) = GenerateOutputParameter("@Updated", adBigInt, 8)
                          
    Dim Updated As Long
    Updated = RunParametricStoredProcedure("Update_Cust_By_NewCarryFee", Parameter)
    If Updated > 0 Then ShowDisMessage " €ÌÌ—«  «‰Ã«„ ‘œ", 1000 Else ShowDisMessage "œ— À»   €ÌÌ—«  „‘ò· ÊÃÊœ œ«—œ", 2000
    DefaultSetting
End Sub

Private Sub cmdUpdatePaykFee_Click()
If Val(txtNewCarryFee.Text) < 0 Or Val(txtOldCarryFee.Text) < 0 Then
        frmMsg.fwlblMsg.Caption = "„ﬁœ«— ﬂ—«ÌÂ ÅÌﬂ ‰„Ì  Ê«‰œ ò„ — «“ ’›— »«‘œ "
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    If txtNewPaykFee.Text = "" Or txtOldPaykFee.Text = "" Then
        frmMsg.fwlblMsg.Caption = "„ﬁœ«— ﬂ—«ÌÂ ÅÌﬂ  Œ«·Ì «”  ·ÿ›« ﬂ—«ÌÂ Õ„· —« Ê«—œ ﬂ‰Ìœ"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
     
    frmMsg.fwlblMsg.Caption = " ¬Ì« „Ì ŒÊ«ÂÌœ ﬂ—«ÌÂ ÅÌﬂ ÃœÌœ Ã«Ìê“Ì‰ ﬂ—«ÌÂ ÅÌﬂ ﬁ»·Ì ê—œœ "
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(0).Caption = "»·Ì"
    frmMsg.fwBtn(1).ButtonType = flwButtonNo
    frmMsg.fwBtn(1).Caption = "ŒÌ—"
    frmMsg.Show vbModal
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
 
     
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@OldPaykFee", adDouble, 8, Val(txtOldPaykFee.Text))
    Parameter(1) = GenerateInputParameter("@NewPaykFee", adDouble, 8, Val(txtNewPaykFee.Text))
    Parameter(2) = GenerateInputParameter("@PercentPaykFee", adDouble, 8, 0)
    Parameter(3) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
    Parameter(4) = GenerateOutputParameter("@Updated", adBigInt, 8)
                          
    Dim Updated As Long
    Updated = RunParametricStoredProcedure("Update_Cust_By_NewPaykFee", Parameter)
    If Updated > 0 Then ShowDisMessage " €ÌÌ—«  «‰Ã«„ ‘œ", 1000 Else ShowDisMessage "œ— À»   €ÌÌ—«  „‘ò· ÊÃÊœ œ«—œ", 2000
    DefaultSetting
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

Private Sub cmdUpdateDistance_Click()
    frmMsg.fwlblMsg.Caption = " ¬Ì« „Ì ŒÊ«ÂÌœ ﬂ—«ÌÂ Â«—« »Â  „«„ „‘ —ﬂÌ‰ «Ì‰ „ÕœÊœÂ  Œ’Ì’ œÂÌœ "
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(0).Caption = "»·Ì"
    frmMsg.fwBtn(1).ButtonType = flwButtonNo
    frmMsg.fwBtn(1).Caption = "ŒÌ—"
    frmMsg.Show vbModal
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
    If txtCarryFee.Text = "" Or txtPaykFee.Text = "" Then
        frmMsg.fwlblMsg.Caption = "„ﬁœ«— ﬂ—«ÌÂ Õ„· Ê ÅÌﬂ —« Œ«·Ì «”  ·ÿ›« „ÕœÊœÂ —« «‰ Œ«» ﬂ‰Ìœ"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    If Val(txtCarryFee.Text) < 0 Or Val(txtPaykFee.Text) < 0 Then
        frmMsg.fwlblMsg.Caption = "„ﬁœ«— ﬂ—«ÌÂ Õ„· Ê ÅÌﬂ ‰„Ì  Ê«‰œ ò„ — «“ ’›— »«‘œ "
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    
     If Trim(txtDescription.Text) = "" Then
            frmMsg.fwlblMsg.Caption = "·ÿ›« „ÕœÊœÂ —« «‰ Œ«» ﬂ‰Ìœ"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
    End If
    
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@Distance", adInteger, 4, txtDescription.Tag)
    Parameter(1) = GenerateInputParameter("@CarryFee", adDouble, 8, Val(txtCarryFee.Text))
    Parameter(2) = GenerateInputParameter("@PaykFee", adDouble, 8, Val(txtPaykFee.Text))
    Parameter(3) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
    Parameter(4) = GenerateOutputParameter("@Updated", adBigInt, 8)
                          
    Dim Updated As Long
    Updated = RunParametricStoredProcedure("Update_Cust_By_Distance", Parameter)
            
End Sub
Private Sub Form_Load()

    If ClsFormAccess.frmDistance = False Then
        Unload Me
        Exit Sub
    End If
    
    CenterCenter Me
    
    VarActForm = Me.Name
    
    With vsDistance
        .Cols = 5
        .TextMatrix(0, 1) = "„ÕœÊœÂ"
        .TextMatrix(0, 2) = "ﬂ—«ÌÂ Õ„·"
        .TextMatrix(0, 3) = "ﬂ—«ÌÂ ÅÌﬂ"

        .ColHidden(4) = True

        .FixedAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignRightCenter

        .ColWidth(0) = 510
        .ColWidth(1) = 1740
        .ColWidth(2) = 1740
        .ColWidth(3) = 1740
    End With

    MyFormAddEditMode = ViewMode
    DefaultSetting
    SetFirstToolBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
     VarActForm = ""
End Sub
Public Sub ExitForm()

    Unload Me
    
End Sub
Public Sub Edit()
    If vsDistance.Rows > 1 Then
        MyFormAddEditMode = EditMode 'Edit
        SetFirstToolBar
    End If
End Sub

Public Sub Delete()

    If vsDistance.Rows < 2 Then Exit Sub

    If MyFormAddEditMode <> 0 Then
        Cancel
    End If
    On Error GoTo ErrHandler
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, txtDescription.Tag)
    RunParametricStoredProcedure "Delete_tblTotal_tDistance_By_Code", Parameter
    
    frmMsg.fwlblMsg.Caption = "»« „Ê›ﬁÌ  Õ–› ‘œ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
    
    DefaultSetting
Exit Sub
    
ErrHandler:
If err.Number = -2147217873 Then

    frmMsg.fwlblMsg.Caption = "„ «”›«‰Â ‘„« ﬁ«œ— »Â Õ–› ‰„Ì »«‘Ìœ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
End If
    
End Sub

Public Sub DefaultSetting()

    Dim Rst As New ADODB.Recordset
    
    Set Rst = RunStoredProcedure2RecordSet("Get_All_tDistance")
    
    With vsDistance
        .Rows = 1
        If Not (Rst.BOF = True And Rst.EOF = True) Then
            While Rst.EOF <> True
                .Rows = .Rows + 1

                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .TextMatrix(.Rows - 1, 1) = Rst!Description
                .TextMatrix(.Rows - 1, 2) = Rst!carryfee
                .TextMatrix(.Rows - 1, 3) = Rst!PaykFee
                .TextMatrix(.Rows - 1, 4) = Rst!Code
                Rst.MoveNext
            Wend
        End If
    
    End With
    
    If Rst.State = 1 Then Rst.Close
     
    Dim Obj As Object
    For Each Obj In Me
        If TypeOf Obj Is TextBox Then
            Obj.Text = ""
            Obj.Tag = 0
        ElseIf TypeOf Obj Is ComboBox Then
            Obj.ListIndex = 0
        ElseIf TypeOf Obj Is OptionButton Then
            Obj.Value = False
        ElseIf TypeOf Obj Is CheckBox Then
            Obj.Value = vbUnchecked
        End If
    Next Obj
    
    Set Rst = Nothing
    
End Sub
Public Sub Add()
    
    MyFormAddEditMode = AddMode
    DefaultSetting
    SetFirstToolBar
    
End Sub

Public Sub Cancel()
   Select Case MyFormAddEditMode
        Case AddMode 'new
            DefaultSetting
            MyFormAddEditMode = AddMode
            SetFirstToolBar
            Add
            
        Case EditMode 'edit
             vsDistance_Click
    End Select
  
End Sub
Public Sub ChangeLanguage()

    Select Case clsStation.Language
    
        Case Farsi
        
        Case English
        
    End Select
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Public Sub Update()
    Dim i As Integer
    ReDim Parameter(3) As Parameter
    Dim Result As Integer
    Dim Obj As Object

    If Trim$(txtDescription.Text) = "" Or Trim$(txtCarryFee.Text) = "" Or Trim$(txtPaykFee.Text) = "" Then
            frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  —« ò«„· Ê«—œ ‰„«ÌÌœ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            
            txtDescription.SetFocus
            
            Exit Sub

    End If
    
    Select Case MyFormAddEditMode
    
        Case AddMode
            Parameter(0) = GenerateInputParameter("@Description", adWChar, 50, Trim(txtDescription.Text))
            Parameter(1) = GenerateInputParameter("@CarryFee", adDouble, 8, Val(Trim(txtCarryFee.Text)))
            Parameter(2) = GenerateInputParameter("@PaykFee", adDouble, 8, Val(Trim(txtPaykFee.Text)))
            Parameter(3) = GenerateOutputParameter("@intStatus", adInteger, 4)
            
            Result = RunParametricStoredProcedure("Insert_tblTotal_tDistance", Parameter)
            
            If Parameter(3).Value <> -1 Then
                txtDescription.Tag = Parameter(3).Value
                frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal

                MyFormAddEditMode = ViewMode
                DefaultSetting
                SetFirstToolBar
            Else
                frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ÃœÌœ À»  ‰‘œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
            End If
            
        Case EditMode
        
            ReDim Parameter(4) As Parameter
            
            Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, txtDescription.Tag)
            Parameter(1) = GenerateInputParameter("@Description", adWChar, 50, Trim(txtDescription.Text))
            Parameter(2) = GenerateInputParameter("@CarryFee", adDouble, 8, Val(Trim(txtCarryFee.Text)))
            Parameter(3) = GenerateInputParameter("@PaykFee", adDouble, 8, Val(Trim(txtPaykFee.Text)))
            Parameter(4) = GenerateOutputParameter("@intStatus", adInteger, 4)
            
            Result = RunParametricStoredProcedure("Update_tblTotal_tDistance", Parameter)

            If Parameter(4).Value <> -1 Then
            
                frmMsg.fwlblMsg.Caption = " €ÌÌ— «ÿ·«⁄«  »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                
                MyFormAddEditMode = ViewMode
                DefaultSetting
                SetFirstToolBar
            Else
            
                frmMsg.fwlblMsg.Caption = "„ «”›«‰Â «ÿ·«⁄«   €ÌÌ— ‰Ì«› . ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                
            End If
            
    End Select

    
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
   If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtCarryFee_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtPaykFee_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub vsDistance_Click()
    
    With vsDistance
        If .Row = 0 Then Exit Sub
        txtDescription.Tag = .TextMatrix(.Row, 4)
        txtDescription.Text = .TextMatrix(.Row, 1)
        txtCarryFee.Text = .TextMatrix(.Row, 2)
        txtPaykFee.Text = .TextMatrix(.Row, 3)

        MyFormAddEditMode = ViewMode
        SetFirstToolBar
    End With
    
End Sub
