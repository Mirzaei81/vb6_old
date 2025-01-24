VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmGroupView 
   BackColor       =   &H00FF8080&
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   FillStyle       =   6  'Cross
   Icon            =   "frmGroupView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   12000
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   0
      Left            =   7080
      TabIndex        =   0
      Tag             =   "frmFacRecursive"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1323
      Caption         =   "·Ì”  ›Ì‘Â«Ì „—ÃÊ⁄Ì"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontBold        =   -1  'True
      FontSize        =   11.25
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   135
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   1
      Left            =   1830
      TabIndex        =   1
      Tag             =   "frmFacEdit"
      Top             =   1815
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1323
      Caption         =   "·Ì”  ›Ì‘Â«Ì «’·«ÕÌ"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontBold        =   -1  'True
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   2
      Left            =   7200
      TabIndex        =   2
      Tag             =   "frmHistory"
      Top             =   4935
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1323
      Caption         =   "”Ê«»ﬁ ›«ò Ê—"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontBold        =   -1  'True
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   3
      Left            =   1830
      TabIndex        =   3
      Tag             =   "frmUserHistory"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1323
      Caption         =   "”Ê«»ﬁ ﬂ«—»—"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontBold        =   -1  'True
      FontSize        =   11.25
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmGroupView.frx":A4C2
      TabIndex        =   5
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‰„«Ì‘ «ÿ·«⁄« "
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
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmGroupView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CounterRep As Integer
Dim Parameter() As Parameter
Dim i As Integer


Private Sub Form_Activate()

    mdifrm.Toolbar3.Visible = False
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

    If ClsFormAccess.frmGroupView = False Then
        Unload Me
        Exit Sub
    End If
    
    CenterTop Me
    
    Dim obj As Object
    For Each obj In Forms
        If TypeOf obj Is Form Then
            If obj.Name <> "mdifrm" And obj.Name <> Me.Name And obj.Name <> "frmAbout" Then
                obj.Hide
            End If
        End If

    Next obj
    
    VarActForm = Me.Name
    
    Dim Rst As New ADODB.Recordset
    Dim strFormName As String
    Dim i As Integer
    
    strFormName = ""
    For i = fwBtnRep.LBound To fwBtnRep.UBound
        If Trim(fwBtnRep(i).Tag) <> "" Then
            strFormName = strFormName & "'" & fwBtnRep(i).Tag & "',"
        End If
    Next i
    
    If strFormName <> "" Then
        strFormName = Mid(strFormName, 1, Len(strFormName) - 1)
    End If
    
    For i = fwBtnRep.LBound To fwBtnRep.UBound
    
        fwBtnRep(i).Enabled = False
        
    Next i
    
    If Rst.State <> 0 Then Rst.Close
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@UserId", adInteger, 4, mvarCurUserNo)
    Parameter(1) = GenerateInputParameter("@intObjectType", adInteger, 4, 1)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set Rst = RunParametricStoredProcedure2Rec("GetUserAccess", Parameter)
        
    For i = 0 To Me.fwBtnRep.Count - 1
        Me.fwBtnRep(i).Enabled = False
    Next i
    If Not (Rst.EOF = True And Rst.BOF = True) Then
            While Rst.EOF <> True
                For i = 0 To Me.fwBtnRep.Count - 1
                    If LCase(Me.fwBtnRep(i).Tag) = LCase(Rst.Fields("ObjectId").Value) Then
                        Me.fwBtnRep(i).Enabled = True
                        Exit For
                    End If
                Next i
                Rst.MoveNext
                
            Wend
    End If

Set Rst = Nothing


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

''''Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''Dim i As Integer
''''For i = 0 To 3
''''    ImageRpt(i).SpecialEffect = fmSpecialEffectBump
''''Next i
''''End Sub



'Private Sub Form_Resize()
'    On Error Resume Next
'     If Me.ScaleHeight > 0 Then
'        Me.Height = iHeight
'        Me.Width = iWidth
'     End If
'End Sub

Public Sub ExitForm()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)

    VarActForm = ""
    
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
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


End Sub


Private Sub fwBtnRep_Click(Index As Integer)

    Unload Me

    Dim obj As Object
    For Each obj In Forms
        If TypeOf obj Is Form Then
            If obj.Name <> "mdifrm" And obj.Name <> "frmAbout" Then
                obj.Hide
            End If
        End If

    Next obj

    Select Case Index
        Case 0
            If ClsFormAccess.frmFacRecursive = True Then
                frmFacRecursive.Show
            End If
        Case 1
            If ClsFormAccess.frmFacEdit = True Then
                frmFacEdit.Show
            End If
        Case 2
            If ClsFormAccess.frmHistory = True Then
                frmHistory.Show
            End If
        Case 3
            'If ClsFormAccess.frmUserHistory = True Then
                frmUserHistory.Show
            'End If
    End Select

End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub Timer1_Timer()
    fwBtnRep(CounterRep).Visible = True
    CounterRep = CounterRep + 1
    If CounterRep = 4 Then
        Timer1.Enabled = False
    End If
End Sub

Private Sub SetFirstToolBar()

    AllButton vbOff, True

    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True

End Sub

