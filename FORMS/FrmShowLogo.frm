VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmShowLogo 
   Caption         =   "                                                                                ÎáÇÕå ÝÇßÊæÑ"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   Icon            =   "FrmShowLogo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   4335
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label lblCarryFeeTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblServiceTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblPackingTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label LblCarryFee 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "˜ÑÇíå Íãá"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label LblDutyTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label lblDuty 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÚæÇÑÖ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label LblService 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÓÑæíÓ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label LblPacking 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÈÓÊå ÈäÏí"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label LblTax 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ãÇáíÇÊ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblTaxTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblDiscountTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblSumPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   3600
         Width           =   2355
      End
      Begin VB.Label LblDiscount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÊÎÝíÝ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÌãÚ  ˜á   "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2760
         TabIndex        =   4
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label LblSubTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÌãÚ ßÇáÇåÇ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.Image Image1 
         Height          =   3855
         Left            =   0
         Picture         =   "FrmShowLogo.frx":A4C2
         Stretch         =   -1  'True
         Top             =   240
         Width           =   4695
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "FrmShowLogo.frx":BF79
      TabIndex        =   8
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmShowLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim p() As Parameter
Dim Flag As Boolean
Dim InputPriceFlag As Boolean

'Private Sub chkInputprice_Click()
'    If chkInputprice.Value = True Then
'        InputPriceFlag = True
'        DoCalculate
'        SaveSetting strMainKey, Me.Name, "chkInputprice", 1
'    Else
'        InputPriceFlag = False
'        SaveSetting strMainKey, Me.Name, "chkInputprice", 0
'    End If
'
'End Sub

Public Sub ClearGridValue()
    
    lblCarryFeeTotal.Caption = ""
    lblDiscountTotal.Caption = ""
    lblPackingTotal.Caption = ""
    lblServiceTotal.Caption = ""
    LblSubTotal.Caption = ""
    lblTaxTotal.Caption = ""
    LblDutyTotal.Caption = ""
    
    lblSumPrice.Caption = ""

End Sub

Public Sub UpdateLblValue()

    lblCarryFeeTotal.Caption = frmInvoice.lblCarryFeeTotal
    lblDiscountTotal.Caption = frmInvoice.lblDiscountTotal
    lblPackingTotal.Caption = frmInvoice.lblPackingTotal
    lblServiceTotal.Caption = frmInvoice.lblServiceTotal
    LblSubTotal.Caption = frmInvoice.LblSubTotal
    lblTaxTotal.Caption = frmInvoice.lblTaxTotal
    LblDutyTotal.Caption = frmInvoice.LblDutyTotal
    
    lblSumPrice.Caption = frmInvoice.lblSumPrice.Caption
       
End Sub


Private Sub Form_Load()
    CenterTop Me
    
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

'    If mVarAccessLevel = 1 Then chkInputprice.Enabled = True Else chkInputprice.Enabled = False
'    If Val(GetSetting(strMainKey, Me.Name, "chkInputprice")) = 0 Then
'        chkInputprice.Value = False
'        InputPriceFlag = False
'    Else
'        InputPriceFlag = True
'        chkInputprice.Value = True
'        FWBtnOK.Enabled = False
'        FWBtnPrint.Enabled = False
'    End If
    
    On Error GoTo ErrHandler
    Dim LogoFile As String
    Dim filetemp As New FileSystemObject
    LogoFile = App.Path & "\Image\Logo.gif"
    If filetemp.FileExists(LogoFile) Then
        Image1.Picture = LoadPicture(LogoFile)
    Else
        LogoFile = App.Path & "\Image\Logo.jpg"
        If filetemp.FileExists(LogoFile) Then
            Image1.Picture = LoadPicture(LogoFile)
        End If
    End If
    
Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 1500
    Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub


'Private Sub TxtPayment_KeyDown(KeyCode As Integer, Shift As Integer)
'   Form_KeyDown KeyCode, Shift
'End Sub

