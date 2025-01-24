VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmShowPictureGood 
   BackColor       =   &H80000016&
   Caption         =   "                     "
   ClientHeight    =   5835
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   6915
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   6915
   Begin VB.Timer TimerShow 
      Left            =   480
      Top             =   0
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmShowPictureGood.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmShowPictureGood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Result As Boolean
Dim i As Integer
Private rctmp As New ADODB.Recordset
Dim No As Double
Dim intSerialNo As Double

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case Shift
      Case 0
          Select Case KeyCode
            Case 13
              SendKeys "{Tab}", 12
                Case 113  ' F2
                       
                  
                 End Select
    End Select
End Sub

Private Sub Form_Load()

    CenterCenter Me
    
    TimerShow.Interval = clsInvoiceValue.ShowGoodTime
    Result = False
      
    GetDataDetail
    formloadFlag = False
    Me.Left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 5000 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 5000 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.Top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    If Me.Left < 0 Then Me.Left = 0
    If Me.Top < 0 Then Me.Top = 0
    If Me.Top > Me.ScaleHeight Then Me.Top = 0

    formloadFlag = True
    
    TimerShow.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub
Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
Sub GetDataDetail()

    Dim TempStr As String
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intCode", adInteger, 4, frmInvoice.GoodCode)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tGood_Picture", Parameter)
    Dim ii As Integer
    If Not (rctmp.BOF Or rctmp.EOF) Then
        
          
         
          
          '' On Error Resume Next
          On Error GoTo err
           Image1.Picture = LoadPicture(rctmp!PicturePath)
    Else
           Image1.Picture = LoadPicture("")
          
    End If
    
err:
 If err.Number = 53 Then
 
        Image1.Picture = LoadPicture("")
        frmMsg.fwlblMsg.Caption = "⁄ﬂ” „Ê—œ ‰Ÿ— Å«ﬂ ‘œÂ «” "
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"

        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.Show vbModal
 End If
    rctmp.Close
    
    End Sub

Private Sub TimerShow_Timer()
    Unload Me
End Sub
