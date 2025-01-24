VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmPhoneBook 
   ClientHeight    =   3840
   ClientLeft      =   300
   ClientTop       =   420
   ClientWidth     =   7995
   Icon            =   "frmPhoneBook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   7995
   Begin VB.Frame Frame1 
      Height          =   2685
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   7755
      Begin VB.TextBox txtEmail 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtAddress 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2160
         Width           =   5895
      End
      Begin VB.TextBox txtFax 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtTelCompany 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtMobile 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtFirstName 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtLastName 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtTel 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         Caption         =   " Å”   «·ﬂ —Ê‰Ìﬂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label lblAddress 
         Alignment       =   1  'Right Justify
         Caption         =   "   ‰‘«‰Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label lblFax 
         Alignment       =   1  'Right Justify
         Caption         =   "   ‰„«»—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblTelCompany 
         Alignment       =   1  'Right Justify
         Caption         =   "    ·›‰ ‘—ﬂ "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblMobile 
         Alignment       =   1  'Right Justify
         Caption         =   "    ·›‰ Â„—«Â"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblFirstName 
         Alignment       =   1  'Right Justify
         Caption         =   "   ‰«„"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblLastName 
         Alignment       =   1  'Right Justify
         Caption         =   "* ‰«„ Œ«‰Ê«œêÌ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblTel 
         Alignment       =   1  'Right Justify
         Caption         =   "*  ·›‰"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   1395
      End
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   525
      Left            =   6480
      Top             =   0
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   32896
      ForeColor2      =   128
      BackColor       =   9412754
      Caption         =   "„—Ê—"
      Alignment       =   2
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmPhoneBook.frx":A4C2
      TabIndex        =   19
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "œ› —çÂ  ·›‰"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   -120
      Width           =   3015
   End
   Begin VB.Label lblFooter 
      Alignment       =   1  'Right Justify
      Caption         =   "*  ò„Ì· «ÿ·«⁄«  »—«Ì ⁄‰«ÊÌ‰ ” «—Â œ«— «Ã»«—Ì „Ì »«‘œ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3480
      Width           =   4605
   End
End
Attribute VB_Name = "frmPhoneBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim cn As New ADODB.Connection
Dim Parameter() As Parameter
Dim intPhoneBookNo As Integer
Public PreviousForm As Form

Public Sub ChangeLanguage()
    
    DefaultSetting
    
End Sub

Public Sub Add()
    
    MyFormAddEditMode = AddMode
    DefaultSetting
    SetFirstToolBar
    
End Sub
Public Sub Find()

        frmFindPhoneBook.Show vbModal
        
        If mvarcode <> 0 Then
            ReDim Parameter(0) As Parameter
            Dim Rst As New ADODB.Recordset
            
            Parameter(0) = GenerateInputParameter("@intPhoneBookNo", adInteger, 4, mvarcode)
            Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_tPhoneBook_Info", Parameter)
            If Rst.State = 1 Then
                If Not (Rst.EOF = True And Rst.BOF = True) Then
                    GetRecrdsetDetail Rst
                End If
            End If
            Set Rst = Nothing
    
            MyFormAddEditMode = ViewMode
            SetFirstToolBar
            mvarcode = 0
            
        Else
            Exit Sub
            
        End If
    
End Sub

Public Sub Printing()

'    ReDim Parameter(0) As Parameter
'
'    Dim MyCrystalReport
'    Set MyCrystalReport = CreateObject("Crystal.CrystalReport")
'
'    Parameter(0) = GenerateInputParameter("@PPNO", adInteger, 4, txtPersonnelNumber.Tag)
'
'    MyCrystalReport.ReportFileName = App.Path & "\Reports" & RepVer & "\RepIdCard.rpt"
'    MyCrystalReport.Destination = crptToPrinter 'crptToWindow ' '
'    MyCrystalReport.ParameterFields(0) = CStr(Parameter(0).Name) & ";" & CStr(Parameter(0).Value) & ";" & "True"
'
'    MyCrystalReport.RetrieveDataFiles
'    MyCrystalReport.Action = 1
'    MyCrystalReport.PageZoom (150)
'
   
End Sub
Public Sub Cancel()

    MyFormAddEditMode = AddMode
    SetFirstToolBar
    DefaultSetting
    
End Sub

Public Sub DefaultSetting()

    Set rctmp = RunStoredProcedure2RecordSet("Get_tblTotal_PhoneBook_New_intPhoneBookNo")
    intPhoneBookNo = rctmp.Fields("intPhoneBookNo").Value
    
    Dim Obj As Object
    For Each Obj In Me
        If TypeOf Obj Is TextBox Then
            Obj.Text = ""
            Obj.Tag = 0
        End If
    Next Obj
    
    
End Sub

Public Sub Edit()
    MyFormAddEditMode = EditMode
    SetFirstToolBar
End Sub

Private Sub GetRecrdsetDetail(tempRst As ADODB.Recordset)

    DefaultSetting
    
    If tempRst.EOF = True And tempRst.BOF = True Then Exit Sub
    
    intPhoneBookNo = tempRst.Fields("intPhoneBookNo").Value
    txtFirstName.Text = IIf(IsNull(tempRst!nvcFirstName), "", tempRst!nvcFirstName)
    txtLastName.Text = tempRst.Fields("nvcLastName").Value
    txtTel.Text = tempRst.Fields("nvcTelCollection").Value
    txtMobile.Text = IIf(IsNull(tempRst!nvcMobile), "", tempRst!nvcMobile)
    txtTelCompany.Text = IIf(IsNull(tempRst!nvcTelCompany), "", tempRst!nvcTelCompany)
    txtFax.Text = IIf(IsNull(tempRst!nvcFax), "", tempRst!nvcFax)
    txtEmail.Text = IIf(IsNull(tempRst!nvcEmail), "", tempRst!nvcEmail)
    TxtAddress.Text = IIf(IsNull(tempRst!nvcAddress), "", tempRst!nvcAddress)
    
End Sub

Public Sub ExitForm()

    Unload Me
End Sub

Public Sub FirstKey()
    ReDim Parameter(1) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@intPhoneBookNo", adInteger, 4, intPhoneBookNo)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, EnumDirection.FirstRecord)
    Set Rst = RunParametricStoredProcedure2Rec("NavigateInPhoneBook", Parameter)
    If Rst.State = 1 Then
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            GetRecrdsetDetail Rst
        End If
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub

Public Sub PreviousKey()
    ReDim Parameter(1) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@intPhoneBookNo", adInteger, 4, intPhoneBookNo)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, EnumDirection.PreviousRecord)
    Set Rst = RunParametricStoredProcedure2Rec("NavigateInPhoneBook", Parameter)
    If Rst.State = 1 Then
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            GetRecrdsetDetail Rst
        End If
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub

Public Sub NextKey()
    ReDim Parameter(1) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@intPhoneBookNo", adInteger, 4, intPhoneBookNo)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, EnumDirection.NextRecord)
    Set Rst = RunParametricStoredProcedure2Rec("NavigateInPhoneBook", Parameter)
    If Rst.State = 1 Then
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            GetRecrdsetDetail Rst
        End If
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub

Public Sub LastKey()
    ReDim Parameter(1) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@intPhoneBookNo", adInteger, 4, intPhoneBookNo)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, EnumDirection.LastRecord)
    Set Rst = RunParametricStoredProcedure2Rec("NavigateInPhoneBook", Parameter)
    If Rst.State = 1 Then
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            GetRecrdsetDetail Rst
        End If
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub


Public Sub SetFirstToolBar()
    
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(1).Enabled = True   'Home
    mdifrm.Toolbar1.Buttons(2).Enabled = True   'PageUp
    mdifrm.Toolbar1.Buttons(3).Enabled = True   'PageDown
    mdifrm.Toolbar1.Buttons(4).Enabled = True   'End
    
    mdifrm.Toolbar1.Buttons(13).Enabled = True   'Find
    mdifrm.Toolbar1.Buttons(15).Enabled = True   'Print
    
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
        Frame1.Enabled = False
        
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        Frame1.Enabled = True

    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        Frame1.Enabled = True
        
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode

End Sub
Public Sub Update()
    
    ReDim Parameter(9) As Parameter
    Dim Result As Integer
    Dim Obj As Object
    
    If Trim(txtTel.Text) = "" Or Trim(txtLastName.Text) = "" Then
            
            frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  ÷—Ê—Ì —« Ê«—œ ‰„«ÌÌœ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
    End If
    
    Select Case MyFormAddEditMode
        Case AddMode
        
            Parameter(0) = GenerateInputParameter("@nvcFirstName", adVarChar, 50, Trim(txtFirstName.Text))
            Parameter(1) = GenerateInputParameter("@nvcLastName", adVarChar, 50, Trim(txtLastName.Text))
            Parameter(2) = GenerateInputParameter("@nvcTelCollection", adVarChar, 1000, Trim(txtTel.Text))
            Parameter(3) = GenerateInputParameter("@intUserNo", adInteger, 4, mvarCurUserNo)
            Parameter(4) = GenerateInputParameter("@nvcMobile", adVarChar, 30, Trim(txtMobile.Text))
            Parameter(5) = GenerateInputParameter("@nvcTelCompany", adVarChar, 30, Trim(txtTelCompany.Text))
            Parameter(6) = GenerateInputParameter("@nvcFax", adVarChar, 30, Trim(txtFax.Text))
            Parameter(7) = GenerateInputParameter("@nvcEmail", adVarChar, 30, Trim(txtEmail.Text))
            Parameter(8) = GenerateInputParameter("@nvcAddress", adVarChar, 100, Trim(TxtAddress.Text))
            Parameter(9) = GenerateOutputParameter("@intPhoneBookNo", adInteger, 4)
            
            Result = RunParametricStoredProcedure("Insert_tblTotal_tPhoneBook", Parameter)
            
            If Result = -1 Then
                frmMsg.fwlblMsg.Caption = "À»  «‰Ã«„ ‰‘œ" + vbCrLf + " ·›‰  ò—«—Ì „Ì »«‘œ"
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                Exit Sub
            Else
                On Error GoTo ErrHandler
            End If
            
            frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
            
            Add
            
        Case EditMode
        
            ReDim Parameter(10) As Parameter
            
            Parameter(0) = GenerateInputParameter("@intPhoneBookNo", adInteger, 50, intPhoneBookNo)
            Parameter(1) = GenerateInputParameter("@nvcFirstName", adVarChar, 50, Trim(txtFirstName.Text))
            Parameter(2) = GenerateInputParameter("@nvcLastName", adVarChar, 50, Trim(txtLastName.Text))
            Parameter(3) = GenerateInputParameter("@nvcTelCollection", adVarChar, 1000, Trim(txtTel.Text))
            Parameter(4) = GenerateInputParameter("@intUserNo", adInteger, 4, mvarCurUserNo)
            Parameter(5) = GenerateInputParameter("@nvcMobile", adVarChar, 30, Trim(txtMobile.Text))
            Parameter(6) = GenerateInputParameter("@nvcTelCompany", adVarChar, 30, Trim(txtTelCompany.Text))
            Parameter(7) = GenerateInputParameter("@nvcFax", adVarChar, 30, Trim(txtFax.Text))
            Parameter(8) = GenerateInputParameter("@nvcEmail", adVarChar, 30, Trim(txtEmail.Text))
            Parameter(9) = GenerateInputParameter("@nvcAddress", adVarChar, 100, Trim(TxtAddress.Text))
            Parameter(10) = GenerateOutputParameter("@Update", adInteger, 4)
            
            Result = RunParametricStoredProcedure("Update_tblTotal_tPhoneBook", Parameter)
            If Result = -1 Then
                frmMsg.fwlblMsg.Caption = "À»  «‰Ã«„ ‰‘œ" + vbCrLf + " ·›‰  ò—«—Ì „Ì »«‘œ"
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                Exit Sub
            Else
                On Error GoTo ErrHandler
            End If
            
            
            frmMsg.fwlblMsg.Caption = " €ÌÌ— «ÿ·«⁄«  »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
            
            Add
    End Select
    Exit Sub
ErrHandler:
    Select Case err.Number
        Case -2147217873
            frmMsg.fwlblMsg.Caption = "À»  «‰Ã«„ ‰‘œ" + vbCrLf + "«ÿ·«⁄«   ò—«—Ì „Ì »«‘œ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
        Case Else
            MsgBox err.Description, vbOKOnly, err.Number
    End Select
    
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
    SetFirstToolBar

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case Shift
        Case 0
            Select Case KeyCode
                
                Case 33
                    NextKey
                Case 34
                    PreviousKey
                Case 35
                    LastKey
                Case 36
                    FirstKey
            End Select
    
    End Select
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

    If ClsFormAccess.frmPhoneBook = False Then
        Unload Me
        Exit Sub
    End If

    CenterCenter Me
    
    VarActForm = Me.Name
    
    MyFormAddEditMode = AddMode
    DefaultSetting
    SetFirstToolBar
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

     VarActForm = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)


    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


''''    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
