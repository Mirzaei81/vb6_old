VERSION 5.00
Begin VB.Form frmGetCustBarcode 
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3780
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetCustBarcode.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   0
      Width           =   3735
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label lblGetCustBarcode 
         Alignment       =   2  'Center
         Caption         =   "ê—› ‰ ﬂœ( »«—ﬂœ) „‘ —Ì"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.PictureBox Frame3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   5115
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   1440
      Width           =   3735
      Begin VB.TextBox txtMaxCustCode 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox txtEnd2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox txtEnd1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox txtStart2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox txtStart1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   2880
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "»— «”«” ﬂ«—«ﬂ —Â«Ì  ‘—Ê⁄ Ê Å«Ì«‰"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox TxtCustBarcode 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "»— «”«” ﬂœ CR  œ— «‰ Â«Ì ﬂœ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "»— «”«”  ⁄œ«œ «—ﬁ«„ «‘ —«ﬂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   " ⁄œ«œ «—ﬁ«„ ﬂœ „‘ —Ì «“—«”  -Õœ«ﬂÀ— 12 "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   15
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "ﬂ«—«ﬂ —Â«Ì Å«Ì«‰"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "ﬂ«—«ﬂ —Â«Ì ‘—Ê⁄"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   3480
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3480
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   " ⁄œ«œ «—ﬁ«„ ﬂœ „‘ —Ì"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmGetCustBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim mvarbarcode As Boolean
Dim strBarCode As String

Private Sub Form_Activate()
    
    Text1.SetFocus

End Sub

Public Sub barcode()
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@ActDeact", adInteger, 4, 0) ' Only Active
    Parameter(1) = GenerateInputParameter("@Code", adBigInt, 8, Val(left(strBarCode, 20)))
    Set Rst = RunParametricStoredProcedure2Rec("Get_Customer_Code", Parameter)
    
    If Rst.EOF <> True Then
        mvarcode = Rst!Code
        mvarName = Rst![Name]
'        If strCategory = "15" And (clsArya.CustomerId = 11 Or clsArya.CustomerId = 12) Then
'            ReDim Parameter(1) As Parameter
'            Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(mvarcode))
'            Parameter(1) = GenerateInputParameter("@Discount", adInteger, 4, Val(Mid(strBarCode, 7, 2)))
'            Set Rst = RunParametricStoredProcedure2Rec("Update_Customer_Discount", Parameter)
'        End If
    Else
        mvarcode = 0
        mvarName = ""
        frmDisMsg.lblMessage.Caption = " «Ì‰ «‘ —«ﬂ »—«Ì „‘ —Ì   ⁄—Ì› ‰‘œÂ "
        frmDisMsg.Timer1.Interval = 2000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
    End If
    Unload Me
End Sub

Private Sub Form_Load()

    If Val(GetSetting(strMainKey, Me.Name, "TxtCustBarcode")) > 0 Then
        TxtCustBarcode = Val(GetSetting(strMainKey, Me.Name, "TxtCustBarcode"))
    Else
         TxtCustBarcode = 4
    End If
    If Val(GetSetting(strMainKey, Me.Name, "txtMaxCustCode")) > 0 Then
        txtMaxCustCode = Val(GetSetting(strMainKey, Me.Name, "txtMaxCustCode"))
    Else
         txtMaxCustCode = 5
    End If
    txtStart1 = GetSetting(strMainKey, Me.Name, "TxtStart1")
    txtStart2 = GetSetting(strMainKey, Me.Name, "txtStart2")
    txtEnd1 = GetSetting(strMainKey, Me.Name, "txtEnd1")
    txtEnd2 = GetSetting(strMainKey, Me.Name, "txtEnd2")
    
    If Val(GetSetting(strMainKey, Me.Name, "Option1CustBarcode")) = 0 Then
        Option1(0).Value = 1
    ElseIf Val(GetSetting(strMainKey, Me.Name, "Option1CustBarcode")) = 1 Then
        Option1(1).Value = 1
    Else
        Option1(2).Value = 1
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    SaveSetting strMainKey, Me.Name, "TxtCustBarcode", CStr(Val(TxtCustBarcode.Text))
    SaveSetting strMainKey, Me.Name, "txtMaxCustCode", CStr(Val(txtMaxCustCode.Text))
    SaveSetting strMainKey, Me.Name, "txtStart1", CStr(txtStart1.Text)
    SaveSetting strMainKey, Me.Name, "txtStart2", CStr(txtStart2.Text)
    SaveSetting strMainKey, Me.Name, "txtEnd1", CStr(txtEnd1.Text)
    SaveSetting strMainKey, Me.Name, "txtEnd2", CStr(txtEnd2.Text)
    If Option1(0).Value = True Then
        SaveSetting strMainKey, Me.Name, "Option1CustBarcode", "0"
    ElseIf Option1(1).Value = True Then
        SaveSetting strMainKey, Me.Name, "Option1CustBarcode", "1"
    Else
        SaveSetting strMainKey, Me.Name, "Option1CustBarcode", "2"
    End If
End Sub


Private Sub Text1_Change()
    If Len(Text1) = TxtCustBarcode And Option1(0).Value = True Then
        strBarCode = Text1
        barcode
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Option1(2).Value = False Then
        If mvarbarcode = True Then
            Select Case KeyCode
                Case 111, 191: '/ Barcode
                    strBarCode = Text1
                    Me.barcode
                    mvarbarcode = False
                    strBarCode = ""
                    Exit Sub
                  
            End Select
'            strBarCode = strBarCode & ChrW(KeyCode)
        Else
        
            Select Case KeyCode
                
                Case 111, 191: '/ Barcode
                
                    mvarbarcode = True
                
            End Select
            
        End If
    End If
    Select Case KeyCode
            Case 13
                If Shift = 0 And Val(Text1) > 0 And Option1(1).Value = True Then
                    strBarCode = (Text1)
                    barcode
                End If
            Case 27  ' Esc
                mvarbarcode = False
                strBarCode = ""
                Unload Me
                Exit Sub
    End Select

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 8 Then
    ElseIf IsNumeric(ChrW(KeyAscii)) = False Then
        If mvarbarcode = True Then
            If Len(txtEnd1) > 0 And Len(txtEnd2) > 0 Then
            Select Case KeyAscii
                Case Asc(txtEnd1), Asc(txtEnd2): '
                    strBarCode = Text1
                    strBarCode = Right(strBarCode, Val(txtMaxCustCode))
                    Me.barcode
                    mvarbarcode = False
                    strBarCode = ""
                    Exit Sub
            End Select
            End If
        Else
             If Len(txtStart1) > 0 And Len(txtStart2) > 0 Then
            Select Case KeyAscii
                Case Asc(txtStart1), Asc(txtStart2):  '/ Barcode
                    mvarbarcode = True
                    Text1 = ""
            End Select
            End If
        End If
        KeyAscii = 0
        'strCardNo = strCardNo & ChrW(KeyCode)
    End If

End Sub

