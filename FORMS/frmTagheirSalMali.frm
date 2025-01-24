VERSION 5.00
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmTagheirSalMali 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5460
   Icon            =   "frmTagheirSalMali.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   5460
   Begin VB.Frame Frame2 
      Caption         =   "”«· „«·Ì œÌ « »Ì” "
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
      Height          =   2415
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
      Begin VB.ComboBox cmbSalMali2 
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin FLWCtrls.FWCoolButton FWBtnOkay2 
         Height          =   615
         Left            =   480
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Traffic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmTagheirSalMali.frx":A4C2
         DownPicture     =   "frmTagheirSalMali.frx":AD9C
         PictureAlign    =   3
         Caption         =   " «ÌÌœ"
         MaskColor       =   -2147483633
         Style           =   2
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "”«· „«·Ì"
         BeginProperty Font 
            Name            =   "B Traffic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "”«·  „«·Ì «Ì‰ «Ì” ê«Â"
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
      Height          =   2415
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
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
         Height          =   510
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin FLWCtrls.FWCoolButton FWBtnOkay 
         Height          =   615
         Left            =   480
         TabIndex        =   3
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Traffic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmTagheirSalMali.frx":B1EE
         DownPicture     =   "frmTagheirSalMali.frx":BAC8
         PictureAlign    =   3
         Caption         =   " «ÌÌœ"
         MaskColor       =   -2147483633
         Style           =   2
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "”«· „«·Ì"
         BeginProperty Font 
            Name            =   "B Traffic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " €ÌÌ— ”«· „«·Ì"
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
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmTagheirSalMali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyFormAddEditMode As EnumAddEditMode
Dim rs As New ADODB.Recordset
Public Sub SetFirstToolBar()
    Dim i As Integer

    AllButton vbOff, True
   
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
        mdifrm.Toolbar1.Buttons(10).Enabled = True 'Delete
                
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
                
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False 'Delete
        
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False 'Delete
        
    End If
    
    ''HeaderLabel Val(MyFormAddEditMode), fwlblMode
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
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Load()
    CenterCenter Me
    
    FillSalMali
    FillSalMali2
    
End Sub
Private Sub FillSalMali()
    cmbSalMali.Clear
    Set rs = RunStoredProcedure2RecordSet("Get_All_tAccountYears")
    Do While rs.EOF = False
        cmbSalMali.AddItem rs!AccountYear
        rs.MoveNext
    Loop
    Dim i As Integer
    For i = 0 To cmbSalMali.ListCount - 1
        cmbSalMali.ListIndex = i
        If AccountYear = cmbSalMali.Text Then
            Exit For
        End If
    Next
    rs.Close

End Sub
Private Sub FillSalMali2()
    Dim DataBaseAccountYear As Integer
    Set rs = RunStoredProcedure2RecordSet("Get_tAccountYears_Active")
    DataBaseAccountYear = rs!AccountYear
    cmbSalMali2.Clear
    Set rs = RunStoredProcedure2RecordSet("Get_All_tAccountYears")
    Do While rs.EOF = False
        cmbSalMali2.AddItem rs!AccountYear
        rs.MoveNext
    Loop
    Dim i As Integer
    For i = 0 To cmbSalMali2.ListCount - 1
        cmbSalMali2.ListIndex = i
        If DataBaseAccountYear = cmbSalMali2.Text Then
            Exit For
        End If
    Next
    rs.Close

End Sub

Public Sub ExitForm()

    Unload Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    VarActForm = ""
End Sub

Private Sub FWBtnOkay_Click()
    If cmbSalMali.ListCount > 0 Then
        AccountYear = Val(cmbSalMali.Text)
        'mdifrm.sbStatusBar.Panels.Item(4).Text = "”«· „«·Ì : " + AccountYear
        SaveSetting strMainKey, "SalMali", "SalMali", AccountYear
        ShowDisMessage " €ÌÌ— ”«· „«·Ì «Ì‰ «Ì” ê«Â «‰Ã«„ ‘œ", 1400
        
 '       Unload Me
    Else
        frmMsg.fwlblMsg.Caption = "ÂÌç ”«· „«·Ì  ⁄—Ì› ‰‘œÂ «” "
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
    End If
End Sub
Private Sub FWBtnOkay2_Click()
    If cmbSalMali2.ListCount > 0 Then
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali2.Text))
        RunParametricStoredProcedure "Update_tAccountYears", Parameter
        ShowDisMessage " €ÌÌ— ”«· „«·Ì œÌ «»Ì” «‰Ã«„ ‘œ", 1400
        
 '       Unload Me
    Else
        frmMsg.fwlblMsg.Caption = "ÂÌç ”«· „«·Ì  ⁄—Ì› ‰‘œÂ «” "
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
    End If
End Sub

