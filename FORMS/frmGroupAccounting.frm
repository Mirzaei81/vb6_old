VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmGroupAccounting 
   BackColor       =   &H00FF8080&
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11430
   FillStyle       =   6  'Cross
   Icon            =   "frmGroupAccounting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   11430
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   135
      Top             =   135
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   1
      Left            =   4440
      TabIndex        =   0
      Tag             =   "frmPayment"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1323
      Caption         =   "Å—œ«Œ  «“ ’‰œÊﬁ"
      BackColor       =   8421631
      ForeColor       =   4194368
      FontName        =   "B Homa"
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   2
      Left            =   1440
      TabIndex        =   1
      Tag             =   "frmRecieved"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1323
      Caption         =   "œ—Ì«›  »Â ’‰œÊﬁ"
      BackColor       =   8421631
      ForeColor       =   4194368
      FontName        =   "B Homa"
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   3
      Left            =   7440
      TabIndex        =   2
      Tag             =   "frmAccCoding"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1323
      Caption         =   "«— »«ÿ »« ﬂœ Â«Ì Õ”«»œ«—Ì"
      BackColor       =   8421631
      ForeColor       =   4194368
      FontName        =   "B Homa"
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   4
      Left            =   4440
      TabIndex        =   3
      Tag             =   "frmCreateSanad"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1323
      Caption         =   " Ê·Ìœ ”‰œ Õ”«»œ«—Ì"
      BackColor       =   8421631
      ForeColor       =   4194368
      FontName        =   "B Homa"
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   735
      Index           =   0
      Left            =   7440
      TabIndex        =   4
      Tag             =   "frmCashClose"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      Caption         =   "»” ‰ Õ”«» ’‰œÊﬁ"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   12
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   120
      OleObjectBlob   =   "frmGroupAccounting.frx":A4C2
      TabIndex        =   6
      Top             =   600
      Width           =   480
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   5
      Left            =   1440
      TabIndex        =   7
      Tag             =   "frmCreateSanad"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1323
      Caption         =   "—“—Ê"
      BackColor       =   8421631
      ForeColor       =   4194368
      FontName        =   "B Homa"
      FontSize        =   12
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Õ”«»œ«—Ì"
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
      Height          =   495
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmGroupAccounting"
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

Private Sub Form_Load()

    If ClsFormAccess.frmGroupAccounting = False Then
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

    Dim Rst As New ADODB.Recordset
    Dim strFormName As String
    Dim i As Integer
    
    strFormName = ""
    
    For i = fwBtnRep.LBound To fwBtnRep.UBound
        If Trim(fwBtnRep(i).Tag) <> "" Then
            strFormName = strFormName & "'" & fwBtnRep(i).Tag & "',"
        End If
    Next i
    
    If strFormName <> "'" Then
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
                If ClsFormAccess.frmCashClose = True And clsArya.Accounting Then
                   frmCashClose.Show
                End If
        Case 1
                If ClsFormAccess.frmPayment = True And clsArya.Accounting Then
                    frmPayment.Show
                End If
        Case 2
                If ClsFormAccess.frmRecieved = True And clsArya.Accounting Then
                    frmRecieved.Show
                End If
        Case 3
                If ClsFormAccess.frmAccCoding = True And clsArya.Accounting Then
                   frmAccCoding.Show
                End If
        
        Case 4
                If ClsFormAccess.frmCreateSanad = True And clsArya.Accounting = True Then
                    frmCreateSanad.Show
                End If
        
    End Select
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    fwBtnRep(CounterRep).Visible = True
    CounterRep = CounterRep + 1
    If CounterRep = 6 Then
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

