VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmGroupStore 
   BackColor       =   &H00FF8080&
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   FillStyle       =   6  'Cross
   Icon            =   "frmGroupStore.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   12000
   Tag             =   "frmGroupStore"
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   720
      OleObjectBlob   =   "frmGroupStore.frx":A4C2
      TabIndex        =   15
      Top             =   240
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   135
      Top             =   135
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   1
      Left            =   7320
      TabIndex        =   0
      Tag             =   "frmGood"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      Caption         =   " ⁄—Ì› ﬂ«·«"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   0
      Left            =   9600
      TabIndex        =   1
      Tag             =   "frmCodingGood"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   "ﬂœÌ‰ê ﬂ«·«"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   5
      Left            =   9240
      TabIndex        =   2
      Tag             =   "frmMojodiControl"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      Caption         =   "ﬂ‰ —·  ﬂ«·«Â«Ì Œ—Ìœ‰Ì Ê ›—ÊŒ ‰Ì"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   9.75
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   9
      Left            =   240
      TabIndex        =   3
      Tag             =   "frmMojodiControl"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1508
      Caption         =   "ê—œ‘ ò«·«Â« œ— «‰»«—"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   12
      Left            =   4680
      TabIndex        =   4
      Tag             =   "frmManageStore"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      Caption         =   "«‰»«— ê—œ«‰Ì"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   10
      Left            =   9360
      TabIndex        =   5
      Tag             =   "frmInventory"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      Caption         =   " ⁄—Ì› «‰»«—"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   2
      Left            =   5040
      TabIndex        =   6
      Tag             =   "frmInventory_Level1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      Caption         =   "«Œ ’«’ ê—ÊÂÂ« »Â «‰»«—"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   3
      Left            =   2760
      TabIndex        =   7
      Tag             =   "frmStation_Inventory"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      Caption         =   "«Œ ’«’ «‰»«— »Â «Ì” ê«Â"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   11
      Left            =   6960
      TabIndex        =   9
      Tag             =   "frmPurchase"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1508
      Caption         =   "Œ—Ìœ - ÕÊ«·Â - —”Ìœ"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   13
      Left            =   2280
      TabIndex        =   10
      Tag             =   "frmCycleStock"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1508
      Caption         =   "«‰»«— ê—œ«‰Ì œÊ—Â «Ì"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   14
      Left            =   120
      TabIndex        =   11
      Tag             =   "frmCustGoodTurnOver"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   "ê—œ‘ ﬂ«·«Ì „‘ —Ì«‰"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Tag             =   "frmUsePercent"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1508
      Caption         =   "÷—Ì» „’—› ﬂ«·« Ê ﬁÌ„   „«„ ‘œÂ"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   9.75
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   6
      Left            =   7080
      TabIndex        =   13
      Tag             =   "frmMojodiControl"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   "ﬂ‰ —·  „ÊÃÊœÌ „Ê«œ «Ê·ÌÂ"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontBold        =   -1  'True
      FontSize        =   9.75
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   8
      Left            =   2760
      TabIndex        =   14
      Tag             =   "frmMojodiControl"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   "ﬂ‰ —· ﬂ«·«Â«Ì ›—ÊŒ ‰Ì"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   11.25
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   7
      Left            =   4920
      TabIndex        =   16
      Tag             =   "frmMojodiControl"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   "ﬂ‰ —· „«‰œÂ ﬂ«·«Â«Ì Ê«”ÿÂ"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "B Homa"
      FontSize        =   9.75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«‰»«—  Ê ﬂ«·« Ê ﬂ‰ —· ¬‰"
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
      TabIndex        =   8
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmGroupStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CounterRep As Integer
Dim Parameter() As Parameter
Dim i As Integer

Private Sub Form_Activate()

    VarActForm = Me.Name
    SetFirstToolBar

''''    Select Case clsArya.ProductSystem
''''        Case False
''''            me.fwBtnRep(2).Enabled = False
''''            me.fwBtnRep(4).Enabled = False
''''            me.fwBtnRep(5).Enabled = False
''''        Case True
''''            me.fwBtnRep(2).Enabled = True
''''            me.fwBtnRep(4).Enabled = True
''''            me.fwBtnRep(5).Enabled = True
''''    End Select
  
  
  End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

    If ClsFormAccess.frmGroupStore = False Then
        Unload Me
        Exit Sub
    End If

    Dim i As Integer
    
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
                        'Exit For
                    End If
                Next i
                Rst.MoveNext
                
            Wend
            If Me.fwBtnRep(5).Enabled = True Then
               Me.fwBtnRep(6).Enabled = True        ' MojodiControl_2
               Me.fwBtnRep(8).Enabled = True        ' MojodiControl_2
               Me.fwBtnRep(7).Enabled = True
            End If
    End If

    Set Rst = Nothing
    Me.fwBtnRep(0).Enabled = True
    Me.fwBtnRep(1).Enabled = True
    Me.fwBtnRep(2).Enabled = True
    Me.fwBtnRep(3).Enabled = True

If clsArya.StoreGroup = False Then
    Me.fwBtnRep(4).Enabled = False
    Me.fwBtnRep(5).Enabled = False
    Me.fwBtnRep(6).Enabled = False
    Me.fwBtnRep(7).Enabled = False
    Me.fwBtnRep(8).Enabled = False
    Me.fwBtnRep(9).Enabled = False
    Me.fwBtnRep(10).Enabled = False
    Me.fwBtnRep(11).Enabled = False
    Me.fwBtnRep(12).Enabled = False
    Me.fwBtnRep(13).Enabled = False
End If

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

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    VarActForm = ""
    
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
    

Select Case Index
    Case 0
        If ClsFormAccess.frmCodingGood = True Then
            frmGood.SSTab1.Tab = 0
            frmGood.Show
        End If
        
    Case 1
        If ClsFormAccess.frmGood = True Then
            frmGood.SSTab1.Tab = 1
            frmGood.Show
        End If
    
    Case 2
        If ClsFormAccess.frmInventory_level1 Then
            frmInventory_level1.Show
        End If
    
    Case 3
        If ClsFormAccess.frmStation_Inventory Then
            frmStation_Inventory.Show
        End If
    Case 4
        If ClsFormAccess.frmUsePercent = True Then
            frmUsePercent.Show
        End If
     
     Case 5
        If ClsFormAccess.frmMojodiControl = True Then
            frmMojodiControl.Show
        End If
    
    Case 6
        If ClsFormAccess.frmMojodiControl = True Then
           frmMojodiControl_2.Show
        End If
        
    Case 7
        If ClsFormAccess.frmRemainingControl = True Then
            frmMojodiControl_3.Show
        End If
    
    Case 8
        If ClsFormAccess.frmRemainingControl = True Then
            frmMojodiControl_4.Show
        End If
   
    Case 9
        If ClsFormAccess.frmRemainingControl = True Then
            frmGoodTurnOver.Show
        End If
    
    Case 10
          If ClsFormAccess.frmInventory Then
            frmInventory.Show
        End If
    
    Case 11
        If ClsFormAccess.frmPurchase Then
            frmPurchase.Show
        End If
    Case 12
        If ClsFormAccess.frmManageStore = True Then
            frmManageStore.Show
        End If
    
   
    Case 13
        If ClsFormAccess.frmCycleStock Then
           frmCycleStock.Show
        End If
    
    
    Case 14
        If ClsFormAccess.frmCustGoodTurnOver Then
           frmCustGoodTurnOver.Show
        End If
        
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
    If CounterRep = 15 Then
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

