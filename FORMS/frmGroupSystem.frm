VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmGroupSystem 
   BackColor       =   &H00FF8080&
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   FillStyle       =   6  'Cross
   Icon            =   "frmGroupSystem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   12000
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   135
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   0
      Left            =   8160
      TabIndex        =   0
      Tag             =   "frmBranch"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1323
      Caption         =   " ⁄—Ì› ‘⁄»Â"
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   735
      Index           =   1
      Left            =   4560
      TabIndex        =   1
      Tag             =   "frmPrinter"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      Caption         =   "Å—Ì‰ —Â«"
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   735
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Tag             =   "frmDeviceSetting"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      Caption         =   "Ê”«Ì· Ã«‰»Ì"
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   3
      Left            =   8400
      TabIndex        =   3
      Tag             =   "frmNotice"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1323
      Caption         =   "‘⁄«—Â«"
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   4
      Left            =   4680
      TabIndex        =   4
      Tag             =   "frmStationsetting"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1323
      Caption         =   "”«Ì—  ‰ŸÌ„« "
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Tag             =   "frmWorkTime"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1323
      Caption         =   " ‰ŸÌ„ ”«⁄  ò«—Ì"
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmGroupSystem.frx":A4C2
      TabIndex        =   6
      Top             =   0
      Width           =   480
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   975
      Index           =   6
      Left            =   8160
      TabIndex        =   7
      Tag             =   "frmStationsetting"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1720
      Caption         =   " ‰ŸÌ„ ›«ﬂ Ê— ›—Ê‘ Ê „Ê‰Ì Ê— œÊ„"
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   7
      Left            =   4560
      TabIndex        =   8
      Tag             =   "frmTaxAndDuty"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1323
      Caption         =   "⁄Ê«—÷ Ê „«·Ì« "
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   8
      Left            =   840
      TabIndex        =   9
      Tag             =   "frmPos"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1323
      Caption         =   " ⁄—Ì› ŒÊœÅ—œ«“ »«‰òÌ"
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   9
      Left            =   8280
      TabIndex        =   10
      Tag             =   "frmTestPos"
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1323
      Caption         =   "  ”  ŒÊœÅ—œ«“ »«‰òÌ"
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   10
      Left            =   4680
      TabIndex        =   11
      Tag             =   "frmTestPos"
      Top             =   6240
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1323
      Caption         =   "—“—Ê"
      Enabled         =   0   'False
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   750
      Index           =   11
      Left            =   840
      TabIndex        =   12
      Tag             =   "frmTestPos"
      Top             =   6240
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1323
      Caption         =   "—“—Ê"
      Enabled         =   0   'False
      BackColor       =   8421631
      ForeColor       =   7362318
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ‰ŸÌ„«   ”Ì” „"
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
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmGroupSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CounterRep As Integer
Dim Parameter() As Parameter
Dim i As Integer
Dim filetemp As New FileSystemObject

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

    If ClsFormAccess.frmGroupSystem = False Then
        Unload Me
        Exit Sub
    End If

    CenterTop Me
    
    CounterRep = 0

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
    
    If intVersion = Min Then
        Me.fwBtnRep(0).Enabled = False
        Me.fwBtnRep(3).Enabled = False
        Me.fwBtnRep(4).Enabled = False
  '      Me.fwBtnRep(8).Enabled = False
    ElseIf intVersion = Normal Then
        Me.fwBtnRep(0).Enabled = False
  '      Me.fwBtnRep(8).Enabled = False
    End If
    
    If fwBtnRep(4).Enabled = True Then fwBtnRep(6).Enabled = True
    Set Rst = Nothing
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

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim i As Integer
'For i = 0 To 3
'    ImageRpt(i).SpecialEffect = fmSpecialEffectBump
'Next i
'End Sub

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
    mdifrm.Toolbar1.Buttons(20).Enabled = False
    mdifrm.Toolbar1.Buttons(21).Enabled = False
 '   mdifrm.Toolbar1.Buttons(27).Enabled = False
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


Private Sub fwBtnRep_Click(index As Integer)
    
    Unload Me
    
    Dim obj As Object
    For Each obj In Forms
        If TypeOf obj Is Form Then
            If obj.Name <> "mdifrm" And obj.Name <> "frmAbout" Then
                obj.Hide
            End If
        End If

    Next obj
    
    Select Case index
        Case 0
            If ClsFormAccess.frmBranch = True Then
                frmBranch.Show
            Else
                Me.Show
            End If
            
        Case 1
            If ClsFormAccess.frmPrinter = True Then
                frmPrinter.Show
            Else
                Me.Show
            End If
        Case 2
            If ClsFormAccess.frmDeviceSetting = True Then
                frmDeviceSetting.Show
            End If
        Case 3
            If ClsFormAccess.frmNotice = True Then
                frmNotice.Show
            Else
                Me.Show
            End If
        Case 4
            If ClsFormAccess.frmStationsetting = True Then
                frmStationsetting.Show
            Else
                Me.Show
            End If
            
        Case 5
        
            If ClsFormAccess.frmWorkTime = True Then
                frmWorkTime.Show
            Else
                Me.Show
            End If
            
        Case 6
        
            If ClsFormAccess.frmStationsetting = True Then
                frmInvoicesetting.Show
            Else
                Me.Show
            End If
        
        Case 7
        
            If ClsFormAccess.frmTaxAndDuty = True Then
                frmTaxAndDuty.Show
            Else
                Me.Show
            End If
        
        Case 8
        
            If ClsFormAccess.frmPos = True Then
                frmPos.Show
            Else
                Me.Show
            End If
        
        Case 9
        
            If clsStation.PosPayment = False Or clsStation.PosModel = 0 Then
                ShowDisMessage "”—ÊÌ” ŒÊœÅ—œ«“ «Ã—« ‰‘œÂ »«Ìœ œ— ›—„ ”«Ì—  ‰ŸÌ„«  „Ê«—œ „—»Êÿ »Â ŒÊœÅ—œ«“ »«‰òÌ «‰ Œ«» ‘Êœ", 2500
                Exit Sub
            End If
            
            If filetemp.FileExists("C:\pcpos\PCPOS.dll") = True Then
            ElseIf filetemp.FileExists(SystemFolderName & "\PCPOS.dll") = True Then
            Else
                ShowDisMessage "›«Ì· C:\pcpos\PCPOS.DLL ÊÃÊœ ‰œ«—œ", 2000
                Exit Sub
            End If
            
            If mdifrm.Winsock_Pos.State <> sckConnected Then mdifrm.PosSocketInit
            If ClsFormAccess.frmTestPos = True Then
                frmTestPos.Show
            Else
                Me.Show
            End If
            
    End Select
    
    
End Sub

Private Sub fwBtnRep_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    KeyActi vbtxtbox, KeyCode, Shift, frmGroupSystem
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
    If CounterRep = fwBtnRep.Count Then
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


