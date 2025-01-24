VERSION 5.00
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{75D4F6FF-8785-11D3-93AD-0000832EF44D}#1.3#0"; "FAST2005.ocx"
Begin VB.Form frmUpdateLock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                        À»  „‘Œ’«  ﬁ›· ”—Ê—"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   4560
      Width           =   4935
      Begin VB.CommandButton cmdCodeRegister 
         BackColor       =   &H0000C0C0&
         Caption         =   "çﬂ ﬂ—œ‰ òœ ﬁ›·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtRegister 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   13
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "À»  „‘Œ’«  œ—Ì«› Ì"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   5880
      Width           =   4935
      Begin VB.CommandButton cmdSave 
         Caption         =   "À»  „‘Œ’«  œ— ﬁ›·"
         Default         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdEscape 
         Caption         =   "Œ—ÊÃ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "œ—ŒÊ«”  «ÿ·«⁄«  «“ ‘—ﬂ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin FLWCtrls.FWNumericTextBox txtPcNo 
         Height          =   525
         Left            =   1560
         TabIndex        =   16
         Top             =   1440
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   926
         Max             =   99
         Min             =   1
         Value           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox CmbVersion 
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox txtSerialNo 
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
         Height          =   405
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin FLWCtrls.FWCheck chkAccounting 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   3720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         Value           =   0   'False
         Caption         =   "Õ”«»œ«—Ì "
         FontName        =   "Tahoma"
         FontSize        =   11.25
         Alignment       =   1
      End
      Begin FLWCtrls.FWNumericTextBox txtPrinterNo 
         Height          =   525
         Left            =   1560
         TabIndex        =   17
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   926
         Max             =   99
         Min             =   1
         Value           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FLWCtrls.FWNumericTextBox txtPpcNo 
         Height          =   525
         Left            =   1560
         TabIndex        =   18
         Top             =   2640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   926
         Max             =   99
         Value           =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblGenerateCodeTag2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ⁄œ«œ ÅÌ œÌ «Ì"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2640
         Width           =   1635
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ⁄œ«œ Å—Ì‰ —Â«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "‰”ŒÂ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   3240
         Width           =   1635
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "”—Ì«· ‰—„ «›“«—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ⁄œ«œ «Ì” ê«ÂÂ«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   1635
      End
   End
   Begin FLWData.FWEncryption FWEncryption1 
      Left            =   0
      Top             =   120
      _ExtentX        =   926
      _ExtentY        =   926
   End
End
Attribute VB_Name = "frmUpdateLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FirstCode, TempCode, LockId, RegisterCode    As String
Dim LenLockid As Long

Private Sub chkAccounting_Click()
    GenerateCode
End Sub

Private Sub CmbVersion_Click()
    GenerateCode
End Sub

Private Sub cmdCodeRegister_Click()
    
    If Len(txtSerialNo) <> 11 Then MsgBox "ÿÊ· «—ﬁ«„ ‘„«—Â ”—Ì«· œ—”  ‰Ì” ": Exit Sub
    
    If TempCode <> Me.txtRegister.Text Then
        MsgBox " „‘Œ’«  œ«œÂ ‘œÂ »«  ‰ŸÌ„«  „‘Œ’ ‘œÂ Â„ ŒÊ«‰Ì ‰œ«—œ"
        cmdSave.Enabled = False
        Frame1.Enabled = True
    Else
        MsgBox "ﬂœ Ê«—œ ‘œÂ çﬂ ‘œ «ﬂ‰Ê‰ „Ì  Ê«‰Ìœ »—‰«„Â —Ì“Ì ﬁ›· —« «œ«„Â œÂÌœ"
        cmdSave.Enabled = True
        Frame1.Enabled = False
    End If
    
End Sub
Private Sub GenerateCode()
    FirstCode = Int(Val(lblGenerateCodeTag2) * 49.543) + 44761
    TempCode = FirstCode
    LockId = (Val(Format(txtPcNo.Value, "00") & Format(txtPrinterNo.Value, "00") & Format(txtPpcNo.Value, "00") & CmbVersion.ItemData(CmbVersion.ListIndex) & IIf(chkAccounting.Value = True, 1, 0)) + Val(FirstCode))
    LockId = Len(LockId) & LockId
    TempCode = Left(TempCode, 3) & CStr(LockId) & Mid(TempCode, 4)
    TempCode = TempCode + Val(txtSerialNo)
End Sub
Private Sub cmdEscape_Click()
    modgl.mvarMsgIdx = vbNo
    Unload Me
End Sub


Private Sub cmdSave_Click()
    strDataLock = txtSerialNo & " = " & CStr(CmbVersion.ItemData(CmbVersion.ListIndex)) & "0" & IIf(chkAccounting.Value = True, 1, 0) & Format(txtPpcNo.Value, "00") & Format(txtPrinterNo.Value, "00") & Format(txtPcNo.Value, "00")
    modgl.mvarMsgIdx = vbYes
    Unload Me
End Sub

Private Sub Form_Activate()
    
    Me.lblGenerateCodeTag2.Caption = Int((Rnd(1)) * 100000)
    Me.lblGenerateCodeTag2.Caption = Int((Rnd(1)) * 100000)

End Sub

Private Sub Form_Load()


    Dim hMenu As Long
    
    hMenu = GetSystemMenu(Me.hwnd, False)

    DeleteMenu hMenu, 6, MF_BYPOSITION
    
    CmbVersion.Clear
    CmbVersion.AddItem "”«œÂ"
    CmbVersion.ItemData(0) = EnumVersion.Min
    CmbVersion.AddItem "„ Ê”ÿ"
    CmbVersion.ItemData(1) = EnumVersion.Normal
    CmbVersion.AddItem " ÅÌ‘—› Â"
    CmbVersion.ItemData(2) = EnumVersion.Silver
    CmbVersion.AddItem "ÊÌéÂ"
    CmbVersion.ItemData(3) = EnumVersion.gold

    cmdSave.Enabled = False
    CmbVersion.ListIndex = 1
    
End Sub

Private Sub txtPc_Changed()
    GenerateCode
End Sub

Private Sub txtPpc_Changed()
    GenerateCode
End Sub

Private Sub txtPrinter_Changed()
    GenerateCode
End Sub

Private Sub txtSerialNo_KeyPress(KeyAscii As Integer)
    If (IsNumeric(ChrW(KeyAscii)) = False Or Len(txtSerialNo) >= 11) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
