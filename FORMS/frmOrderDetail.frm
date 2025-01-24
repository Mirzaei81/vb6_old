VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{687FF23F-0E9B-449D-A782-2C5E7116EDB0}#1.0#0"; "Fardate.ocx"
Begin VB.Form frmOrderDetail 
   BackColor       =   &H80000016&
   Caption         =   "                                        ⁄ÌÌ‰ “„«‰  ÕÊÌ·               "
   ClientHeight    =   2325
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   5055
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
   Icon            =   "frmOrderDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2325
   ScaleWidth      =   5055
   Begin VB.ComboBox cmbShift 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "frmOrderDetail.frx":A4C2
      Left            =   1680
      List            =   "frmOrderDetail.frx":A4C4
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmOrderDetail.frx":A4C6
      TabIndex        =   5
      Top             =   0
      Width           =   480
   End
   Begin MSMask.MaskEdBox mskTime 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   873
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##"
      PromptChar      =   " "
   End
   Begin FLWCtrls.FWButton FWBtnOK 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      Caption         =   "F2 À»   "
      FontBold        =   -1  'True
      Alignment       =   1
   End
   Begin FarDate1.FarDate FarDate1 
      Height          =   345
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Object.ToolTipText     =   "FarDate"
   End
   Begin FLWCtrls.FWButton FWBtnEsc 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      ButtonType      =   1
      Caption         =   "F3 «‰’—«›"
      FontBold        =   -1  'True
      Alignment       =   1
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ  ÕÊÌ·"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”«⁄   ÕÊÌ·"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‘Ì›   ÕÊÌ·"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "frmOrderDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Result As Boolean
Dim i As Integer
 
Private rctmp As New ADODB.Recordset
Dim No As Double
Dim intSerialNo As Long
Dim CountOrder As Integer
Dim clsDate As New clsDate
 
Private Sub cmbShift_GotFocus()
    Sendkey "{F4}", 10
End Sub

Private Sub Form_Activate()
'     FarDate1.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case Shift
      Case 0
          Select Case KeyCode
            Case 13
                SendKeys "{Tab}", 12
            Case 113  ' F2
                FWBtnOK_Click
            Case 114  ' F3
                FwbtnEsc_Click
        End Select
    End Select
End Sub

Private Sub Form_Load()

    CenterCenter Me
    
    Result = False
    
    cmbShift.Clear
'    ReDim Parameter(0) As Parameter
'    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
'    Set rctmp = RunParametricStoredProcedure2Rec("Get_All_tShift", Parameter)
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_tShift")
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            cmbShift.AddItem rctmp!Description
            cmbShift.ItemData(cmbShift.NewIndex) = rctmp!Code
            rctmp.MoveNext
        Wend
'        For i = 1 To cmbShift.ListCount
'            If cmbShift.ItemData(i - 1) = clsStation.NextDeliveryShift Then
'               Me.cmbShift.ListIndex = i - 1
'               Exit For
'            End If
'        Next i
    Else
        cmbShift.AddItem " "
        cmbShift.ItemData(0) = -1
    End If
'    Me.cmbShift.ListIndex = -1
    If rctmp.State = adStateOpen Then If rctmp.State = adStateOpen Then rctmp.Close
'    mskTime.Text = IIf(clsStation.NextDeliveryTime = "", "17:00", clsStation.NextDeliveryTime)
'    FarDate1.Text = clsDate.shamsiAddedDate(Date, Val(IIf(clsStation.NextDeliveryDate = 0, 1, clsStation.NextDeliveryDate)))
    mskTime.Text = "17:00"
    
    FarDate1.Text = "13 " & clsDate.shamsi(Date)

    GetDataDetail
    formloadFlag = False
    Me.left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 5000 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 5000 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    If Me.left < 0 Then Me.left = 0
    If Me.top < 0 Then Me.top = 0
    If Me.top > Me.ScaleHeight Then Me.top = 0

    formloadFlag = True

End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top

    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub

Private Sub FwbtnEsc_Click()
    Unload Me
End Sub

Private Sub FWBtnOK_Click()
If FarDate1.Text = "  /  /  " Or mskTime.Text = "  :  " Or cmbShift.ListIndex = -1 Then
    frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  —« ﬂ«„· Ê«—œ ﬂ‰Ìœ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
    If FarDate1.Text = "  /  /  " Then
        FarDate1.SetFocus
    ElseIf mskTime = "  :  " Then
        mskTime.SetFocus
    Else
        cmbShift.SetFocus
    End If
Else

    If frmInvoice.lblCustomer.Tag <> "-1" Then
        ReDim Parameter(4) As Parameter
        Parameter(0) = GenerateInputParameter("@intSerialNo", adDouble, 8, intSerialNo)
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Parameter(2) = GenerateInputParameter("@Date", adVarChar, 50, Mid(FarDate1.Text, 3))
        Parameter(3) = GenerateInputParameter("@ShiftNo", adInteger, 4, cmbShift.ItemData(cmbShift.ListIndex))
        Parameter(4) = GenerateInputParameter("@Code", adInteger, 4, Val(frmInvoice.lblCustomer.Tag))
        Set rctmp = RunParametricStoredProcedure2Rec("Get_Order_By_Date", Parameter)
        CountOrder = rctmp!CountOrder
    
        If CountOrder > 0 Then
            frmMsg.fwlblMsg.Caption = "»—«Ì «Ì‰ „‘ —Ì œ— «Ì‰  «—ÌŒ Ê ‘Ì›  ﬁ»·« ”›«—‘ À»  ‘œÂ «” "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
'            Exit Sub
        End If
    End If
    Dim ShamsiDateName As String
    ShamsiDateName = clsDate.Miladi(FarDate1.Text)
    ShamsiDateName = clsDate.Find_DayOfWeek(Weekday(ShamsiDateName, vbSaturday))
    ReDim Parameter(5) As Parameter
    Parameter(0) = GenerateInputParameter("@intSerialNo", adDouble, 8, intSerialNo)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(2) = GenerateInputParameter("@Date", adVarChar, 50, Mid(FarDate1.Text, 3))
    Parameter(3) = GenerateInputParameter("@Time", adVarChar, 50, mskTime.Text)
    Parameter(4) = GenerateInputParameter("@ShiftNo", adInteger, 4, cmbShift.ItemData(cmbShift.ListIndex))
    Parameter(5) = GenerateInputParameter("@DayName", adVarChar, 10, ShamsiDateName)
    RunParametricStoredProcedure "Insert_tblTotal_Order", Parameter
    ShowDisMessage "”›«—‘ ÃœÌœ À»  ‘œ", 1000
    Unload Me
End If


End Sub

Private Sub mskTime_GotFocus()
    mskTime.SelStart = 0
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
Sub GetDataDetail()

    If OrderNo = -1 Then Exit Sub
    
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, OrderNo)
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, EnumFactorType.Order)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set rctmp = RunParametricStoredProcedure2Rec("Get_FacM_Per", Parameter)
    
    intSerialNo = rctmp!intSerialNo
  
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@IntserialNo", adBigInt, 8, intSerialNo)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tblTotal_Order", Parameter)
    If Not (rctmp.BOF Or rctmp.EOF) Then
        
        FarDate1.Text = "13" & rctmp!Date
        mskTime.Text = rctmp!time
        cmbShift.ListIndex = rctmp!ShiftNo - 1
               
    End If
    If rctmp.State = adStateOpen Then If rctmp.State = adStateOpen Then rctmp.Close
  
End Sub


