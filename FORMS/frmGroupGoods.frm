VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmGroupGoods 
   BackColor       =   &H00FF8080&
   ClientHeight    =   6465
   ClientLeft      =   105
   ClientTop       =   495
   ClientWidth     =   11820
   FillStyle       =   6  'Cross
   Icon            =   "frmGroupGoods.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   11820
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   135
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   11
      Left            =   7080
      TabIndex        =   0
      Tag             =   "frmGoodDifferences"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   " €ÌÌ—«  œ— ò«·«"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   12
      Left            =   4920
      TabIndex        =   1
      Tag             =   "frmGoodPrintFormats"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      Caption         =   "ò«·« œ— ç«Å"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   7
      Left            =   4920
      TabIndex        =   2
      Tag             =   "frmTable"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   " ⁄—Ì› „Ì“"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   2
      Left            =   4920
      TabIndex        =   3
      Tag             =   "frmPartition"
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   " ⁄—Ì› »Œ‘"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   16
      Left            =   6960
      TabIndex        =   4
      Tag             =   "frmPocketPcGroupsAndStations"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      Caption         =   "Pocket PC  ⁄—Ì› ê—ÊÂ"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   17
      Left            =   4440
      TabIndex        =   5
      Tag             =   "frmPocketPCGroupGood"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1508
      Caption         =   "Pocket PC  ⁄—Ì› ò«·« œ—"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Tag             =   "frmBank"
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   " ⁄—Ì› »«‰ò"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   5
      Left            =   9240
      TabIndex        =   7
      Tag             =   "frmSalMali"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   " ⁄—Ì› ”«· „«·Ì"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   0
      Left            =   9240
      TabIndex        =   8
      Tag             =   "frmKB"
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   "«Œ ’«’ ﬂ«·« »Â ﬂÌ»—œ"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   1
      Left            =   7080
      TabIndex        =   9
      Tag             =   "frmMenu"
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   "«Œ ’«’ ﬂ«·« »Â „‰ÊÂ«"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   13
      Left            =   2640
      TabIndex        =   10
      Tag             =   "frmGoodInKitchen"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   "‰„«Ì‘ ò«·« œ— ¬‘Å“Œ«‰Â"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   10
      Left            =   9240
      TabIndex        =   11
      Tag             =   "frmGoodDiscount"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1508
      Caption         =   " Œ›Ì› ⁄Ê«—÷ „«·Ì«   —ÊÌ ﬂ«·«Â«"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   9.75
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   3
      Left            =   2640
      TabIndex        =   12
      Tag             =   "frmCredit"
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   " ⁄—Ì› »‰"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   6
      Left            =   7080
      TabIndex        =   14
      Tag             =   "frmTagheirSalMali"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   " €ÌÌ— ”«· „«·Ì"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   15
      Left            =   9240
      TabIndex        =   15
      Tag             =   "frmWeeding"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   " Œ’Ì’ „—«”„"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   9
      Left            =   360
      TabIndex        =   16
      Tag             =   "frmPrizeType"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   " ⁄—Ì› ‰Ê⁄ Ã«Ì“Â"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   14
      Left            =   360
      TabIndex        =   17
      Tag             =   "frmAutoDiscount"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   " Œ›Ì› « Ê„« Ìò"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   8
      Left            =   2640
      TabIndex        =   18
      Tag             =   "frmDistance"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   " ⁄—Ì› „ÕœÊœÂ"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmGroupGoods.frx":A4C2
      TabIndex        =   19
      Top             =   0
      Width           =   480
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   18
      Left            =   2400
      TabIndex        =   20
      Tag             =   "frmManageSet"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1508
      Caption         =   " ⁄«—Ì› Å«ÌÂ"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin FLWCtrls.FWRealButton fwBtnRep 
      Height          =   855
      Index           =   19
      Left            =   360
      TabIndex        =   21
      Tag             =   "frmGoodPicture"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      Caption         =   "«Œ ’«’  ’ÊÌ— »Â ﬂ«·«"
      BackColor       =   8421631
      ForeColor       =   4194304
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   12
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„‰ÊÌ  ⁄«—Ì›"
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
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmGroupGoods"
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

    If ClsFormAccess.frmGroupGoods = False Then
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
                        If (LCase(Rst.Fields("ObjectId").Value) = LCase("frmGoodInKitchen") Or LCase(Rst.Fields("ObjectId").Value) = LCase("frmTable") Or LCase(Rst.Fields("ObjectId").Value) = LCase("frmPocketPCGroupGood") Or LCase(Rst.Fields("ObjectId").Value) = LCase("frmPocketPcGroupsAndStations") Or LCase(Rst.Fields("ObjectId").Value) = LCase("frmPartition") Or LCase(Rst.Fields("ObjectId").Value) = LCase("frmGoodDifferences")) And (Not (Val(strCategory) >= 0 And Val(strCategory) <= 7)) Then
                            Me.fwBtnRep(i).Enabled = False
                            Me.fwBtnRep(i).Caption = ""
                        End If
                        If LCase(Rst.Fields("ObjectId").Value) = LCase("frmGoodInKitchen") And ((Val(strCategory) >= 0 And Val(strCategory) <= 7)) And clsArya.MaxStationNo < 1 Then
                            Me.fwBtnRep(i).Enabled = False
                          Me.fwBtnRep(i).Caption = ""
                        End If
                        If (LCase(Rst.Fields("ObjectId").Value) = LCase("frmPocketPcGroupsAndStations") Or LCase(Rst.Fields("ObjectId").Value) = LCase("frmPocketPCGroupGood")) And clsArya.MaxPocketPcNo < 1 Then
                            Me.fwBtnRep(i).Enabled = False
                            Me.fwBtnRep(i).Caption = ""
                       End If
                        Exit For
                    End If
                Next i
                Rst.MoveNext
                
            Wend
    End If

Set Rst = Nothing

CounterRep = 0
If clsArya.TableGarson = False Then
    Me.fwBtnRep(7).Enabled = False
End If
If clsArya.MaxPocketPcNo = 0 Then
    Me.fwBtnRep(16).Enabled = False
    Me.fwBtnRep(17).Enabled = False
End If
If clsArya.MaxKitchenNo = 0 Then
    Me.fwBtnRep(13).Enabled = False
End If
If clsArya.Delivery = False Then
    Me.fwBtnRep(8).Enabled = False
End If
    
    If intVersion = Min Then
        Me.fwBtnRep(7).Enabled = False
        Me.fwBtnRep(8).Enabled = False
        Me.fwBtnRep(9).Enabled = False
        Me.fwBtnRep(10).Enabled = False
        Me.fwBtnRep(11).Enabled = False
        Me.fwBtnRep(12).Enabled = False
        Me.fwBtnRep(13).Enabled = False
  '      Me.fwBtnRep(14).Enabled = False
        Me.fwBtnRep(15).Enabled = False
        Me.fwBtnRep(16).Enabled = False
        Me.fwBtnRep(17).Enabled = False
    ElseIf intVersion = Normal Then
        Me.fwBtnRep(9).Enabled = False
        Me.fwBtnRep(10).Enabled = False
'        Me.fwBtnRep(11).Enabled = False
'        Me.fwBtnRep(12).Enabled = False
        Me.fwBtnRep(13).Enabled = False
 '       Me.fwBtnRep(14).Enabled = False
        Me.fwBtnRep(15).Enabled = False
        Me.fwBtnRep(16).Enabled = False
        Me.fwBtnRep(17).Enabled = False
    ElseIf intVersion = Silver Then
        Me.fwBtnRep(15).Enabled = False
        Me.fwBtnRep(16).Enabled = False
        Me.fwBtnRep(17).Enabled = False
    End If
    
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
    If clsStation.TreeViewMenu = False Then mdifrm.Toolbar3.Visible = True
    
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
        If ClsFormAccess.frmKB = True Then
            frmKB.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
        
    Case 1
    
        If ClsFormAccess.frmMenu = True Then
            frmMenu.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
        
    Case 2
    
        If ClsFormAccess.frmTable = True Then
            frmPartition.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
        
    Case 3
        If ClsFormAccess.frmCredit = True Then
            frmCredit.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    Case 4
        If ClsFormAccess.frmBank = True Then
            frmBank.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    Case 5
        If ClsFormAccess.frmSalMali = True Then
            frmSalMali.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    Case 6
        If ClsFormAccess.frmTagheirSalMali = True Then
            frmTagheirSalMali.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    Case 7
    
        If ClsFormAccess.frmTable = True Then
            frmTable.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
                
    Case 8
        If clsArya.SurroundPayk = True Then
            If ClsFormAccess.frmDistance = True Then
                frmDistance.Show
            End If
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    Case 9
        If ClsFormAccess.frmPrizeType = True Then
            frmPrizeType.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    
    Case 10
    
        If ClsFormAccess.frmGoodDiscount = True Then
            frmGoodDiscount.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    Case 11
        If ClsFormAccess.frmGoodDifferences = True Then
            frmGoodDifferences.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    Case 12
        If ClsFormAccess.frmGoodPrintFormats = True Then
            frmGoodPrintFormats.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    Case 13
    
        If ClsFormAccess.frmGoodInKitchen = True Then
            frmGoodInKitchen.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    
    Case 14
        If ClsFormAccess.frmAutodiscount = True Then
            frmAutodiscount.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    Case 15
        If ClsFormAccess.frmWeeding = True Then
            frmWeeding.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    Case 16
    
        If ClsFormAccess.frmPocketPcGroupsAndStations = True Then
            frmPocketPcGroupsAndStations.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    
    Case 17
    
        If ClsFormAccess.frmPocketPCGroupGood = True Then
            frmPocketPCGroupGood.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    Case 18
        If ClsFormAccess.frmManageSet = True Then
            frmManageSet.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
    Case 19
        If ClsFormAccess.frmGoodPicture = True Then
            frmGoodPicture.Show
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
End Select

End Sub

Private Sub fwBtnRep_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    KeyActi vbtxtbox, KeyCode, Shift, frmGroupGoods
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
    If CounterRep = 20 Then
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

