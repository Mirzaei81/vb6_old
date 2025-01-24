VERSION 5.00
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form FrmMsgTransport 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   Icon            =   "FrmMsgTransport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FLWCtrls.FWButton FWBtnYes 
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2520
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      Caption         =   "»·Ì"
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
      Alignment       =   1
   End
   Begin FLWCtrls.FWButton FWBtnNo 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "«‰’—«› Ê Œ—ÊÃ"
      Top             =   2520
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      ButtonType      =   1
      Caption         =   "ŒÌ—"
      BackColor       =   12632256
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   14.25
      Alignment       =   1
      Object.ToolTipText     =   "«‰’—«› Ê Œ—ÊÃ"
   End
   Begin VB.Label LblNext 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   465
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label lblCurrent 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   465
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "¬Ì« ⁄„·Ì«  «‰ ﬁ«· „ÊÃÊœÌ «“ œÊ—Â Ã«—Ì »Â œÊ—Â »⁄œÌ «‰Ã«„ ‘œÂ Ê œÊ—Â Ã«—Ì ﬁ›· ‘Êœø"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   6345
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "‰«„ œÊ—Â »⁄œÌ:"
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
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "‰«„ œÊ—Â Ã«—Ì:"
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
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1305
   End
End
Attribute VB_Name = "FrmMsgTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim Parameter() As Parameter
Dim TmpBranch, TmpCycleStockNo As Integer

Private Sub Form_Load()
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, TmpBranch)
    Parameter(1) = GenerateInputParameter("@CycleStockNo", adInteger, 4, TmpCycleStockNo)
    Set rs = RunParametricStoredProcedure2Rec("Get_tblTotal_Inventory_Good_CycleStock_Name", Parameter)
    lblCurrent = rs!FirstUnlockCycleStockName
    LblNext = rs!NextUnlockCycleStockName
    If rs!NextUnlockCycleStockName = "" Then
        FWBtnYes.Enabled = False
    End If
End Sub
Public Function SendVariables(ByRef Branch, ByRef CycleStockNo)
    TmpBranch = Branch
    TmpCycleStockNo = CycleStockNo
End Function

Private Sub Form_Unload(Cancel As Integer)
 If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub

Private Sub FWBtnNo_Click()
    mvarIndexNo = 0
    Unload Me
End Sub

Private Sub FWBtnYes_Click()
    mvarIndexNo = 1
    Unload Me
End Sub
