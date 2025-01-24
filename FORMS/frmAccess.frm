VERSION 5.00
Begin VB.Form frmAccess 
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3030
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   3030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   " «∆Ìœ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«‰’—«›"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox TxtUserName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "  —„“ „œÌ— Ì« œ” —”Ì »«·«‰— —« Ê«—œ ﬂ‰Ìœ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1215
      Left            =   120
      Top             =   1320
      Width           =   2835
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ﬂ·„Â ⁄»Ê—"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "‰«„ ò«—»—"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1440
      Width           =   1035
   End
End
Attribute VB_Name = "frmAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ModeAccess As String
Public MyFormAddEditMode As EnumAddEditMode
Public ReturnAccess As Boolean
Public AccessStatus As EnumAccessStatus


Private Sub Command1_Click(Index As Integer)
    On Error GoTo Err_Handler
    
    Dim Rst As New ADODB.Recordset
    Dim CurUserName As Integer
    ReturnAccess = False
'    If Index = 1 Then 'cancel
'        Unload frmAccess
'        Exit Sub
'    End If
            
    If (Index = 1 And AccessAfterClosingcash = False) Then 'cancel and it is not after closing the cash
        Unload frmAccess
        Exit Sub
    ElseIf (Index = 1 And AccessAfterClosingcash = True) Then 'cancel and it's after closing the cash
        Exit Sub
    End If
        
        Set Rst = modgl.GetPerInfo(TxtUserName.Text, txtPassword.Text, CurrentBranch)
        
        If Rst.EOF = True And Rst.BOF = True Then
                       
            frmMsg.fwlblMsg.Caption = "ﬂ·„Â ⁄»Ê— œ—”  ‰Ì” "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Me.txtPassword.Text = ""
            Me.txtPassword.SetFocus
            Exit Sub
        ElseIf Rst.Fields("ActDeAct") = 0 Then
           frmMsg.fwlblMsg.Caption = "ò«—»— €Ì— ›⁄«· «”  "
           frmMsg.fwBtn(0).Visible = False
           frmMsg.fwBtn(1).ButtonType = flwButtonOk
           frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
           frmMsg.Show vbModal
           Me.txtPassword.Text = ""
           Me.txtPassword.SetFocus
           Exit Sub
        ElseIf Rst.Fields("Job") = 9 Then
           frmMsg.fwlblMsg.Caption = "ê«—”Ê‰ ‰„Ì  Ê«‰œ Ê«—œ ”Ì” „ ‘Êœ"
           frmMsg.fwBtn(0).Visible = False
           frmMsg.fwBtn(1).ButtonType = flwButtonOk
           frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
           frmMsg.Show vbModal
           Me.txtPassword.Text = ""
           Me.txtPassword.SetFocus
           Exit Sub
        Else
            
            If AccessAfterClosingcash = False And AccessStatus <> LockShow Then

                CurUserName = Rst.Fields("UID")
                Rst.Cancel
                
                If Rst.State <> 0 Then Rst.Close
                ReDim Parameter(2) As Parameter
                Parameter(0) = GenerateInputParameter("@UserId", adInteger, 4, CurUserName)
                Parameter(1) = GenerateInputParameter("@intObjectType", adInteger, 4, 2)
                Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                
                Set Rst = RunParametricStoredProcedure2Rec("GetUserAccess", Parameter)
            
                 If AccessStatus = Edit Then
                     Select Case MyFormAddEditMode
                     Case RefferedMode
                         ModeAccess = "RefferedAllStationsFactors"
                     Case EditMode, ManipulateMode
                         ModeAccess = "EditAllStationsFactors"
                     End Select
                 ElseIf AccessStatus = CashClose Then
                     ModeAccess = "EditInvoiceCashClose"
                 ElseIf AccessStatus = UpperAmountGood Then
                     ModeAccess = "UpperAmountGood"
                 End If
                 
                 If Not (Rst.EOF = True And Rst.BOF = True) Then
                    Do While Rst.EOF <> True
                        If LCase(ModeAccess) = LCase(Rst.Fields("ObjectId").Value) Then
                            ReturnAccess = True
                            Exit Do
                        End If
                        Rst.MoveNext
                    Loop
                End If
            ElseIf AccessStatus = LockShow Then
                mVarAccessLevel = Rst.Fields("intAccessLevel")
                ReturnAccess = True
            Else 'AccessAfterClosingcash = True
                mvarCurrentLoggedInUserName = Trim(Me.TxtUserName.Text)
                mvarCurUserNo = Rst.Fields("UID")
                mvarPPNo = Rst.Fields("pPno")
                mVarAccessLevel = Rst.Fields("intAccessLevel")
                mvarCountRePrint = Rst.Fields("CountRePrint")
                mvarCountInvoicePrint = Rst.Fields("CountInvoicePrint")
                         
                Dim aa, bb As String
                aa = Rst.Fields("Description")
                bb = Rst.Fields("nvcFirstName") + " " + Rst.Fields("nvcSurName")
                mdifrm.StatusBar1.Panels(4).Text = " ”„  " & " = " & aa
                mdifrm.StatusBar1.Panels(5).Text = "ò«—»—" & " = " & bb
'                ReDim Parameter(2) As Parameter
'                Parameter(0) = GenerateInputParameter("@UserId", adInteger, 4, mvarCurUserNo)
'                Parameter(1) = GenerateInputParameter("@intObjectType", adInteger, 4, 1)
'                Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
'
'                Set Rst = RunParametricStoredProcedure2Rec("GetUserAccess", Parameter)
'                Dim i As Integer
'                Dim Obj As Object
'                For Each Obj In frmGroupMenu
'            '        Debug.Print obj.Name
'                    If TypeOf Obj Is FWCoolButton Then
'                        Obj.Enabled = False
'                    End If
'                Next
'                If Not (Rst.EOF = True And Rst.BOF = True) Then
'                    While Rst.EOF <> True
'                        For Each Obj In frmGroupMenu
'            '                Debug.Print obj.Name
'                            If TypeOf Obj Is FWCoolButton Then
'                                If Obj.Tag = Val(Rst.Fields("intObjectCode").Value) Then
'                                    Obj.Enabled = True
'                                    Exit For
'                                End If
'                            End If
'                        Next
'                        Rst.MoveNext
'                    Wend
'                End If
                ClsFormAccess.Class_Initialize
                ReturnAccess = True
            
            End If
        End If
    
        
'Exit Sub  No Need Because we need return from this form

Err_Handler:
    Dim frmActive As Form
    Dim strFormName As String
    Set Rst = Nothing
    If ReturnAccess = False Then
        frmMsg.fwlblMsg.Caption = "œ” —”Ì ﬂ«›Ì ‰Ì” "
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Me.txtPassword.Text = ""
    Else
        Unload frmAccess
        If AccessStatus = LockShow Then
            AccessStatus = EnumAccessStatus.None
        Else
            If AccessAfterClosingcash = True Then
                Unload frmGroupMenu
                frmGroupMenu.Show  'Reload with new access
                For Each frmActive In Forms
                    If frmActive.Name = "frmInvoice" Then
                        frmInvoice.Show
                        Exit For
                    End If
                Next
            End If
        End If
        
    End If
    
Exit Sub
End Sub
        
Private Sub Form_Activate()
    If TxtUserName.Text <> "" Then
        txtPassword.SetFocus
    Else
        TxtUserName.SetFocus
    End If
    SetKbLayout LANG_EN_US
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then Command1_Click 0
End Sub

Private Sub txtPassword_Validate(Cancel As Boolean)
        txtPassword.TabIndex = 1
        TxtUserName.TabIndex = 0
End Sub



