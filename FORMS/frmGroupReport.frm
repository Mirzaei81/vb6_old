VERSION 5.00
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmGroupReport 
   BackColor       =   &H00FF8080&
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   FillStyle       =   6  'Cross
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   7335
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   870
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Tag             =   "frmCust"
      Top             =   2040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1535
      Caption         =   "ê“«—‘«  „ÊÃÊœ"
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontBold        =   -1  'True
      FontSize        =   14.25
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   120
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   870
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Tag             =   "frmSupplier"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1535
      Caption         =   " Ê·Ìœ ê“«—‘"
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontBold        =   -1  'True
      FontSize        =   14.25
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ê“«—‘« "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "frmGroupReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    VarActForm = ""
    modgl.RightButton True
    mdifrm.Toolbar1.Buttons(20).Enabled = False
    mdifrm.Toolbar1.Buttons(21).Enabled = False
 '   mdifrm.Toolbar1.Buttons(27).Enabled = False

    If clsStation.TreeViewMenu = False Then mdifrm.Toolbar3.Visible = True
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    Dim obj As Object
    For Each obj In Forms
        If TypeOf obj Is Form Then
            If obj.Name <> "mdifrm" And obj.Name <> Me.Name And obj.Name <> "frmAbout" Then
                obj.Show
            End If
        End If

    Next obj
    SaveSetting strMainKey, Me.Name, "Width", Me.Width
    SaveSetting strMainKey, Me.Name, "Height", Me.Height
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
End Sub

Private Sub fwBtnRep_Click(index As Integer)
    Select Case index
    Case 0:
        frmReportsItem.Show
        Unload Me
    Case 1:
        frmReportGenerator.Show
        Unload Me
    End Select
End Sub
Private Sub SetFirstToolBar()

    AllButton vbOff, True

    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True

End Sub
Private Sub Form_Activate()
    mdifrm.Toolbar3.Visible = False
    VarActForm = Me.Name
    SetFirstToolBar

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Load()

    If ClsFormAccess.frmReports = False Then
        Unload Me
        Exit Sub
    End If
    
    CenterTop Me
    VarActForm = Me.Name
'    SetFirstToolBar
'
    Dim obj As Object
    For Each obj In Forms
        If TypeOf obj Is Form Then
            If obj.Name <> "mdifrm" And obj.Name <> Me.Name And obj.Name <> "frmAbout" Then
                obj.Hide
            End If
        End If

    Next obj

'
'Dim Rst As New ADODB.Recordset
'Dim strFormName As String
'Dim i As Integer
'
'strFormName = ""
'For i = fwBtnRep.LBound To fwBtnRep.UBound
'    If Trim(fwBtnRep(i).Tag) <> "" Then
'        strFormName = strFormName & "'" & fwBtnRep(i).Tag & "',"
'    End If
'Next i
'
'If strFormName <> "" Then
'    strFormName = Mid(strFormName, 1, Len(strFormName) - 1)
'End If
'
'For i = fwBtnRep.LBound To fwBtnRep.UBound
'
'    fwBtnRep(i).Enabled = False
'
'Next i
'
'    If Rst.State <> 0 Then Rst.Close
'    ReDim Parameter(2) As Parameter
'    Parameter(0) = GenerateInputParameter("@UserId", adInteger, 4, mvarCurUserNo)
'    Parameter(1) = GenerateInputParameter("@intObjectType", adInteger, 4, 1)
'    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
'
'    Set Rst = RunParametricStoredProcedure2Rec("GetUserAccess", Parameter)
'
'    For i = 0 To Me.fwBtnRep.Count - 1
'        Me.fwBtnRep(i).Enabled = False
'    Next i
'    If Not (Rst.EOF = True And Rst.BOF = True) Then
'            While Rst.EOF <> True
'                For i = 0 To Me.fwBtnRep.Count - 1
'                    If LCase(Me.fwBtnRep(i).Tag) = LCase(Rst.Fields("ObjectId").Value) Then
'                        Me.fwBtnRep(i).Enabled = True
'                        Exit For
'                    End If
'                Next i
'                Rst.MoveNext
'
'            Wend
'    End If

    CounterRep = 0
    formloadFlag = False
    Me.Left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 0 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 0 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.Top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    formloadFlag = True



End Sub
Public Sub ExitForm()
    Unload Me
End Sub
