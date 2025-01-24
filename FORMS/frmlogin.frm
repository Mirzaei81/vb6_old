VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{75D4F148-8785-11D3-93AD-0000832EF44D}#4.0#0"; "FAST2003.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8760
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "«‰’—«›"
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
      Index           =   1
      Left            =   4110
      TabIndex        =   21
      Top             =   2850
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ê—Êœ"
      Default         =   -1  'True
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
      Index           =   0
      Left            =   5910
      TabIndex        =   20
      Top             =   2850
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
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
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   4110
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   1230
      Width           =   1815
   End
   Begin VB.TextBox TxtUserName 
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
      Height          =   420
      Left            =   4110
      TabIndex        =   12
      Top             =   750
      Width           =   1815
   End
   Begin VB.CommandButton cmd_FgaryaSite 
      Caption         =   "www.Fgarya.com"
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
      Left            =   4140
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1980
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox frame3 
      BackColor       =   &H00FFFF80&
      Height          =   4335
      Left            =   5490
      RightToLeft     =   -1  'True
      ScaleHeight     =   4275
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   1170
      Width           =   3735
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Text            =   "00-000-00"
         Top             =   2550
         Width           =   1815
      End
      Begin VB.ComboBox CmbVersion 
         BackColor       =   &H00FFFFFF&
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
         Left            =   390
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3630
         Width           =   1755
      End
      Begin VB.CommandButton CmdSpec 
         Caption         =   "„‘Œ’« "
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
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2970
         Width           =   1725
      End
      Begin VB.ComboBox cmbStation 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1020
         Width           =   1815
      End
      Begin VB.ComboBox CmbSalMali 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1980
         Width           =   1815
      End
      Begin VB.ComboBox cmbBranch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1500
         Width           =   1815
      End
      Begin VB.ComboBox cmbDataBase 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TxtDataSource 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   150
         TabIndex        =   1
         Text            =   "."
         Top             =   30
         Width           =   1815
      End
      Begin VB.Label serial 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”—Ì«·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2310
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   3060
         Width           =   1155
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰”ŒÂ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   3750
         Width           =   615
      End
      Begin VB.Label LblSerialNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   3150
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„‘Œ’«  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   2580
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ì” ê«Â"
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
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   990
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "”«· „«·Ì"
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
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1950
         Width           =   1635
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ ‘⁄»Â"
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
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1470
         Width           =   1635
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»«‰ﬂ  «ÿ·«⁄« Ì"
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
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ ”—Ê—"
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
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   0
         Width           =   1635
      End
   End
   Begin FLWCtrls.FWToolTip FWToolTip 
      Left            =   0
      Top             =   840
      _ExtentX        =   926
      _ExtentY        =   926
      ForeColor       =   -2147483625
      BackColor       =   65535
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin FLWSystem.FWSysInfo FWSysInfo1 
      Left            =   0
      Top             =   1320
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmlogin.frx":9670
      TabIndex        =   11
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Date 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   585
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   120
      Width           =   2955
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ò«—»—"
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
      Left            =   6030
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   750
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ·„Â ⁄»Ê—"
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
      Left            =   6030
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   1230
      Width           =   1635
   End
   Begin VB.Image Image_Fgarya 
      Height          =   4725
      Left            =   -210
      Picture         =   "frmlogin.frx":96F6
      Stretch         =   -1  'True
      Top             =   -900
      Width           =   4815
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsDate As New clsDate
Private Counter As String
Private DatabaseVersion As Integer
Private SoftwareVersion As Integer
Private CurrentScriptNo As Integer
Dim rctmp As New ADODB.Recordset
Dim i As Integer
Dim SalMaliFocusFlag As Boolean
Dim BranchFocusFlag, StationFocusFlag As Boolean
Dim f As New FileSystemObject
Dim filetemp As New FileSystemObject
Dim tempstring As TextStream
Dim Str As String
Dim IsFileExist As Boolean
Dim Rst As New ADODB.Recordset
Dim strtemporary As String
Dim strtemporary1 As String
Dim strtemporary2 As String
Dim Result As Integer
Dim txtUserIsFocus As Boolean
Dim strPolicy As String
Private WithEvents TinyEvent As TINYLib.Tiny
Attribute TinyEvent.VB_VarHelpID = -1


Private Sub cmd_FgaryaSite_Click()
    Dim st As String
    If strDelegate = "56" Then st = "www.MoeinReklam.com" Else st = "www.fgarya.com"
    
    ShellExecute Me.hwnd, vbNullString, st, vbNullString, "C:\", 2   ' SW_SHOWNORMAL
End Sub

Private Sub cmdCapability_Click()
    frmCapability.Show
End Sub

Private Sub CmdSpec_Click()
    
    AccessAfterClosingcash = False
    Unload frmAccess
    frmAccess.lblTitle.Caption = " —„“ „œÌ— —« »—«Ì ‰„«Ì‘ «ÿ·«⁄«  Ê«—œ ﬂ‰Ìœ"
    frmAccess.AccessStatus = LockShow
    frmAccess.Show vbModal
    
    If mVarAccessLevel <> 1 Then
        ShowDisMessage "›ﬁÿ »« œ” —”Ì „œÌ— „Ì  Ê«‰ Ê«—œ «Ì‰ ﬁ”„  ‘œ", 1500
    Else
        Unload frmClientLock
        Load frmClientLock
        
        frmClientLock.txtSerialNo = clsArya.HardLockSerialNo
        frmClientLock.txtPcNo = clsArya.MaxStationNo
        frmClientLock.txtPrinterNo = clsArya.MaxprinterNo
        frmClientLock.txtPpcNo = clsArya.MaxPocketPcNo
        frmClientLock.txtTabletNo = clsArya.MaxTabletNo
        frmClientLock.txtAccNo = clsArya.MaxAccountingNo
        For i = 0 To frmClientLock.CmbVersion.ListCount
            If frmClientLock.CmbVersion.ItemData(i) = intVersion Then
               frmClientLock.CmbVersion.ListIndex = i
               Exit For
            End If
        Next i
        frmClientLock.Show vbModal
    End If
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub Timer1_Timer()
    CloseWindow "On-Screen Keyboard"
    End
End Sub

Private Sub TinyEvent_TinyHIDDidconnect()
    On Error Resume Next
    MsgBox ("ﬁ›· «“ œ” ê«Â Ãœ« ‘œ")
'    ShowDisMessage " ﬁ›· «“ œ” ê«Â Ãœ« ‘œ . ", 2000
    Timer1.Enabled = True  ' For End
End Sub

Public Sub ExitForm()

    Unload Me
    
End Sub
Private Sub SetTooltipText()
    
    With FWToolTip
        .BackColor = vbYellow
        .Ballon = True
        .Margin(flwToolTipMarginLeft) = 20
        .MaxWidth = 500
        .DelayTime(flwToolTipDelayDefault) = 100
        '.DelayTime(flwToolTipDelayInitial) = 100
        .DelayTime(flwToolTipDelayShow) = 4000
        .DelayTime(flwToolTipDelayReshow) = 1500
        .Text(CmbVersion) = "»« «‰ Œ«» Ê—é‰ Â«Ì „Œ ·› ‰—„ «›“«— „Ì Ê«‰Ìœ «“ ﬁ«»·Ì  Â«Ì ¬‰ „ÿ·⁄ ‘ÊÌœ"
        .Text(Label7) = "»« «‰ Œ«» Ê—é‰ Â«Ì „Œ ·› ‰—„ «›“«— „Ì Ê«‰Ìœ «“ ﬁ«»·Ì  Â«Ì ¬‰ „ÿ·⁄ ‘ÊÌœ"
        .Text(LblLimited) = " ‰„«Ì‘ „‘Œ’«  ‰—„ «›“«— ‰’» ‘œÂ "
        .Text(Frame2) = " ‰„«Ì‘ „‘Œ’«  ‰—„ «›“«— ‰’» ‘œÂ "
        .Text(TxtUserName) = "‰«„ ò«—»—Ì —« Ê«—œ ò‰Ìœ."
        .Text(txtPassword) = "—„“ ⁄»Ê— —« Ê«—œ ò‰Ìœ.  "
        .Text(cmd_FgaryaSite) = " «ÿ·«⁄ «“ ¬Œ—Ì‰ „Õ’Ê·«  Ê œ” «Ê—œÂ«Ì «› ÃÌ ¬—Ì« œ— ”«Ì  "
        .Text(Frame1) = strPolicy
    End With
End Sub
Private Sub SetTooltipText_Demo()
    
    With FWToolTip
        .BackColor = vbYellow
        .Ballon = True
        .Margin(flwToolTipMarginLeft) = 20
        .MaxWidth = 500
        .DelayTime(flwToolTipDelayDefault) = 100
        '.DelayTime(flwToolTipDelayInitial) = 100
        .DelayTime(flwToolTipDelayShow) = 4000
        .DelayTime(flwToolTipDelayReshow) = 1500
        If clsArya.LimitedVersion = True And HardLockFlagTrial = False Then
            .Text(CmbVersion) = "»« «‰ Œ«» Ê—é‰ Â«Ì „Œ ·› ‰—„ «›“«— „Ì Ê«‰Ìœ «“ ﬁ«»·Ì  Â«Ì ¬‰ „ÿ·⁄ ‘ÊÌœ"
            .Text(Label7) = "»« «‰ Œ«» Ê—é‰ Â«Ì „Œ ·› ‰—„ «›“«— „Ì Ê«‰Ìœ «“ ﬁ«»·Ì  Â«Ì ¬‰ „ÿ·⁄ ‘ÊÌœ"
            .Text(LblLimited) = "«Ì‰ ‘—ò  œ— ’Ê—  «” ›«œÂ «“ ”Ì” „ ¬“„«Ì‘Ì »—«Ì ‰êÂœ«—Ì œ«œÂ Â« Ê Å‘ Ì»«‰Ì ”Ì” „ ÂÌçêÊ‰Â „”∆Ê·Ì Ì ‰œ«—œ " & vbLf & "œ— ’Ê—  Œ—Ìœ ‰—„ «›“«— œ«œÂ Â« œ— ”Ì” „ «’·Ì ﬁ«»· «” ›«œÂ „Ì »«‘‰œ"
            .Text(Frame2) = "«Ì‰ ‘—ò  œ— ’Ê—  «” ›«œÂ «“ ”Ì” „ ¬“„«Ì‘Ì »—«Ì ‰êÂœ«—Ì œ«œÂ Â« Ê Å‘ Ì»«‰Ì ”Ì” „ ÂÌçêÊ‰Â „”∆Ê·Ì Ì ‰œ«—œ " & vbLf & "œ— ’Ê—  Œ—Ìœ ‰—„ «›“«— œ«œÂ Â« œ— ”Ì” „ «’·Ì ﬁ«»· «” ›«œÂ „Ì »«‘‰œ"
            .Text(TxtUserName) = "‰«„ ò«—»—Ì —« Ê«—œ ò‰Ìœ. „ﬁœ«— ÅÌ‘ ›—÷  " & "0" & "  „Ì »«‘œ"
            .Text(txtPassword) = "—„“ ⁄»Ê— —« Ê«—œ ò‰Ìœ. „ﬁœ«— ÅÌ‘ ›—÷  " & "100" & "  „Ì »«‘œ"
        End If
        .Text(Frame1) = strPolicy
        .Text(cmd_FgaryaSite) = " «ÿ·«⁄ «“ ¬Œ—Ì‰ „Õ’Ê·«  Ê œ” «Ê—œÂ«Ì «› ÃÌ ¬—Ì« œ— ”«Ì  "
    End With
End Sub
''''Private Sub BtnKeypad_Click(Index As Integer)
''''    If txtUserIsFocus = True Then
''''        If Index = 10 Then
''''            If Len(TxtUserName.Text) > 0 Then
''''                TxtUserName.Text = Left(TxtUserName.Text, Len(TxtUserName.Text) - 1)
''''            End If
''''        Else
''''            TxtUserName.Text = TxtUserName.Text & BtnKeypad(Index).Tag
''''        End If
''''    Else
''''        If Index = 10 Then
''''            If Len(txtPassword.Text) > 0 Then
''''                txtPassword.Text = Left(txtPassword.Text, Len(txtPassword.Text) - 1)
''''            End If
''''        Else
''''            txtPassword.Text = txtPassword.Text & BtnKeypad(Index).Tag
''''        End If
''''    End If
''''End Sub

''''Private Sub BtnKeypad_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
''''    If KeyCode = 8 And Shift = 0 Then
''''        BtnKeypad_Click (10)
''''    Else
''''        Command1_Click (0)
''''    End If
''''End Sub
Private Sub FillCmbStation()
    Dim Rst As New ADODB.Recordset
    
    Set Rst = RunStoredProcedure2RecordSet("Get_PC_Stations")
    Dim i As Integer
    i = 0
    cmbStation.Clear
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Not Rst.EOF
            i = i + 1
            cmbStation.AddItem Rst.Fields("Description").Value
            cmbStation.ItemData(cmbStation.ListCount - 1) = Rst.Fields("StationID").Value
            Rst.MoveNext
        Wend
    End If
    
    For i = 0 To cmbStation.ListCount - 1
        If clsArya.StationNo = cmbStation.ItemData(i) Then
            cmbStation.ListIndex = i
            Exit For
        End If
    Next
    If PosConnection.State = 1 Then PosConnection.Close
    Set PosConnection = Nothing
    If Rst.State = 1 Then Rst.Close
    Set Rst = Nothing

End Sub

Private Sub cmbBranch_Change()
'    cmbBranch_Click
End Sub

Private Sub cmbBranch_Click()
    If BranchFocusFlag = False Then Exit Sub
    
    cmbStation_Click
    
'    CurrentBranch = cmbBranch.ItemData(cmbBranch.ListIndex)
    CurrentBranchName = cmbBranch.Text
End Sub

Private Sub cmbBranch_GotFocus()
    BranchFocusFlag = True
End Sub

Private Sub cmbSalMali_Change()
    If SalMaliFocusFlag = False Then Exit Sub
    AccountYear = Val(cmbSalMali.Text)
'    mdifrm.sbStatusBar.Panels.Item(4).Text = "”«· „«·Ì : " + AccountYear
    SaveSetting strMainKey, "SalMali", "SalMali", AccountYear
End Sub

Private Sub cmbSalMali_Click()
    cmbSalMali_Change
End Sub

Private Sub CmbSalMali_GotFocus()
    SalMaliFocusFlag = True
End Sub

Private Sub cmbStation_Click()
    If StationFocusFlag = False Then Exit Sub
    clsArya.StationNo = cmbStation.ItemData(cmbStation.ListIndex)
    StationSettingFile = App.Path & "\Setting\Station" & clsArya.StationNo & ".txt"
    IsFileExist = filetemp.FileExists(StationSettingFile)
    
    If IsFileExist = False Then
      SetDefaultStationSettingFile
      MsgBox "Station Setting File Did Not Exist" & vbCrLf & "Default Station Setting File Created"
    
    End If
    
    clsStation.Class_Initialize
    
    ReDim Parameters(0) As Parameter
    Parameters(0) = GenerateInputParameter("@intStationNo", adInteger, 4, clsArya.StationNo)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_StationIP_ByStationNo", Parameters)
    If rctmp.EOF = False Then
        CurrentBranch = rctmp!Branch
    End If
    For i = 0 To cmbBranch.ListCount - 1
        If CurrentBranch = cmbBranch.ItemData(i) Then
            cmbBranch.ListIndex = i
            Exit For
        End If
    Next
    CurrentBranchName = cmbBranch.Text

End Sub

Private Sub cmbStation_GotFocus()
    StationFocusFlag = True
End Sub

Private Sub CmbVersion_Click()
    intVersion = CmbVersion.ItemData(CmbVersion.ListIndex)
End Sub

Private Sub CmdOnScreenKeyboard_Click()
       ' Shell App.Path & "\Tools\osk.exe"
        Shell SystemFolderName & "\osk.exe"
End Sub

Private Sub Command1_Click(index As Integer)

    If index = 1 Then 'cancel
        If HardLockFlag = True Or HardLockFlagTrial = True Then
            Tiny1.UserPassWord (KarbarKey)
            Tiny1.SetAutoCheckingTinyHID (False)
            Tiny1.DisconnectFromTinyHID
            Timer1.Enabled = True
            Exit Sub
        Else
            Unload Me
            End
        End If
    ElseIf index = 0 Then 'ok
        
        Dim Updated As Boolean
         strtemporary = "|ÅÇhgke"
         strtemporary = DText((strtemporary), frmfactor.Label1.Caption)
         strtemporary1 = "|ÅÇÄhgkf"
         strtemporary1 = DText((strtemporary1), frmfactor.Label1.Caption)
         strtemporary2 = "|ÅÇÅhgkg"
         strtemporary2 = DText((strtemporary2), frmfactor.Label1.Caption)
        
        If Me.txtPassword.Text = strtemporary Then   'Zap Databse
            
              frmMsg.fwlblMsg.Caption = "¬Ì« »—«Ì Å«ﬂ ﬂ—œ‰ œÌ « ›—Ê‘ «ÿ„Ì‰«‰ œ«—Ìœ"
            ' frmMsg.fwBtn(0).Visible = True
              frmMsg.fwBtn(0).ButtonType = flwButtonOk
              frmMsg.fwBtn(1).ButtonType = flwButtonCancel
              frmMsg.fwBtn(0).Caption = "»·Ì"
              frmMsg.fwBtn(1).Caption = "ŒÌ—"
              frmMsg.Show vbModal
              If mvarMsgIdx = vbYes Then
                 ReDim Parameter(0) As Parameter
                 
                 Parameter(0) = GenerateOutputParameter("@Result", adInteger, 4)
            
                 Updated = RunParametricStoredProcedure("Zap_Sale_DataBase", Parameter)
                 If Updated = True Then
                 
                     MsgBox "Å«ﬂ ﬂ—œ‰ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ" & vbLf & "»—‰«„Â —« „Ãœœ« «Ã—« ﬂ‰Ìœ"
                     End
                 Else
                     MsgBox "«‘ﬂ«· œ— Å«ﬂ ”«“Ì œÌ «" & vbLf & "»« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ"
                 End If
             Else
                 MsgBox "»—‰«„Â —« „Ãœœ« «Ã—« ﬂ‰Ìœ"
                 End
             End If
        ElseIf Me.txtPassword.Text = strtemporary1 Then   'Zap Databse
            
              frmMsg.fwlblMsg.Caption = " ¬Ì« »—«Ì Å«ﬂ ﬂ—œ‰ œÌ « ﬂ«·«Â« «ÿ„Ì‰«‰ œ«—Ìœ - œÌ « ›—Ê‘ ‰Ì“ Å«ﬂ ŒÊ«Âœ ‘œ "
            ' frmMsg.fwBtn(0).Visible = True
              frmMsg.fwBtn(0).ButtonType = flwButtonOk
              frmMsg.fwBtn(1).ButtonType = flwButtonCancel
              frmMsg.fwBtn(0).Caption = "»·Ì"
              frmMsg.fwBtn(1).Caption = "ŒÌ—"
              frmMsg.Show vbModal
              If mvarMsgIdx = vbYes Then
                 ReDim Parameter(0) As Parameter
                 
                 Parameter(0) = GenerateOutputParameter("@Result", adInteger, 4)
            
                 Updated = RunParametricStoredProcedure("Zap_Sale_DataBase", Parameter)
                 If Updated = True Then
                     
                     ReDim Parameter(0) As Parameter
                     Parameter(0) = GenerateOutputParameter("@Result", adInteger, 4)
                     Updated = RunParametricStoredProcedure("Zap_Goods_DataBase", Parameter)
                     If Updated = 1 Then
                        MsgBox "Å«ﬂ ﬂ—œ‰ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ" & vbLf & "»—‰«„Â —« „Ãœœ« «Ã—« ﬂ‰Ìœ"
                        End
                     Else
                        MsgBox "«‘ﬂ«· œ— Å«ﬂ ”«“Ì œÌ «Ì ﬂ«·«Â«" & vbLf & "»« ‘—ﬂ  «› ÃÌ ¬—Ì«   „«” »êÌ—Ìœ"
                        End
                     End If
                 Else
                     MsgBox "«‘ﬂ«· œ— Å«ﬂ ”«“Ì œÌ «Ì ›—Ê‘" & vbLf & "»« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ"
                     End
                 End If
             Else
                 MsgBox "»—‰«„Â —« „Ãœœ« «Ã—« ﬂ‰Ìœ"
                 End
             End If
            
        ElseIf Me.txtPassword.Text = strtemporary2 Then   'Zap Databse
            
              frmMsg.fwlblMsg.Caption = "¬Ì« »—«Ì Å«ﬂ ﬂ—œ‰ ﬂ· œÌ «Â« «ÿ„Ì‰«‰ œ«—Ìœ"
            ' frmMsg.fwBtn(0).Visible = True
              frmMsg.fwBtn(0).ButtonType = flwButtonOk
              frmMsg.fwBtn(1).ButtonType = flwButtonCancel
              frmMsg.fwBtn(0).Caption = "»·Ì"
              frmMsg.fwBtn(1).Caption = "ŒÌ—"
              frmMsg.Show vbModal
              If mvarMsgIdx = vbYes Then
                 ReDim Parameter(0) As Parameter
                 
                 Parameter(0) = GenerateOutputParameter("@Result", adInteger, 4)
            
                 Updated = RunParametricStoredProcedure("Zap_All_DataBase", Parameter)
                 If Updated = True Then
                 
                     MsgBox "Å«ﬂ ﬂ—œ‰ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ" & vbLf & "»—‰«„Â —« „Ãœœ« «Ã—« ﬂ‰Ìœ"
                     End
                 Else
                     MsgBox "«‘ﬂ«· œ— Å«ﬂ ”«“Ì œÌ «" & vbLf & "»« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ"
                 End If
             Else
                 MsgBox "»—‰«„Â —« „Ãœœ« «Ã—« ﬂ‰Ìœ"
                 End
             End If
            
         Else
                                                ' Check Database Version
            '------------------------------------------------------
            Dim LastScriptNo As String
            
            ReDim Parameters(0) As Parameter

            Parameters(0) = GenerateInputParameter("@Version", adInteger, 4, SoftwareVersion)

            Set Rst = RunParametricStoredProcedure2Rec("Get_ScriptNo_DataBase", Parameters)
            
            If Rst.EOF = True And Rst.BOF = True Then
                MsgBox "«‘ﬂ«· œ— ŒÊ«‰œ‰ ‘„«—Â ‰”ŒÂ œÌ « »Ì” " & vbLf & "»« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ"
                End
            
            Else
                If Rst!ScriptNo < CurrentScriptNo Then
                    MsgBox "CurrentScriptNo = " & CurrentScriptNo & vbLf & "LastScriptNo = " & Rst!ScriptNo & vbLf & "«Ì‰ ‰”ŒÂ «“ »—‰«„Â »« ‘„«—Â ¬Œ—Ì‰ «”ﬂ—ÌÅ  „—»ÊÿÂ Â„«Â‰ê ‰Ì”  - " & vbLf & "œﬁ  ‘Êœ «„ﬂ«‰ Œÿ« ÊÃÊœ œ«—œ"
                End If
            End If
            
            '------------------------------------------------------
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
'                If CurrentDateNumber > LastDateNumber Then
'                    ShellExecute Me.hwnd, vbNullString, "www.fgarya.com", vbNullString, "C:\", 2   ' SW_SHOWNORMAL
'                End If
                mvarCurrentLoggedInUserName = Trim(Me.TxtUserName.Text)
                mvarCurUserNo = Rst.Fields("UID")
                mvarPPNo = Rst.Fields("pPno")
                mVarAccessLevel = Rst.Fields("intAccessLevel")
                mvarCountRePrint = Rst.Fields("CountRePrint")
                mvarCountInvoicePrint = Rst.Fields("CountInvoicePrint")
                mvarTafsili = IIf(IsNull(Rst.Fields("Tafsili")), 0, Rst.Fields("Tafsili"))
                
                Dim aa, bb As String
                aa = Rst.Fields("Description")
                bb = Rst.Fields("nvcFirstName") + " " + Rst.Fields("nvcSurName")
                
                
                Rst.Cancel
                
                Dim DataBaseName As String
                DataBaseName = cmbDataBase.Text
                

                Unload FrmLogin
                DoEvents
                frmAbout.Show
                mdifrm.Show
                
                'Sleep 1500
                Call ODBCSetting(clsArya.ServerName, DataBaseName)
                
                If clsStation.Language = Farsi Then
                    If strDelegate = "17" Then
                        mdifrm.StatusBar1.Panels(1).Text = " ‘⁄»Â :" & CurrentBranchName & " | " & "»Â ”›«—‘ „Ê””Â Õ”«»—”Ì Œ«ﬁ«‰Ì -  03412521758" & " | " & "«Ã—« :  ‘—ò  ›‰ ¬Ê—ê” —¬—Ì« _  ·›‰ :" & "  +982188554455,+982188554466,+982188554477,+982188554488"
                    ElseIf strDelegate = "24" Then
                        mdifrm.StatusBar1.Panels(1).Text = "»Â ”›«—‘ ‘—ò   òÌ‰ «·ò —Ê‰Ìò Å«”«—ê«œ -  ·›‰ 22263035(09821) ›ò”  22916851" & " | "  ' & "«Ã—« :  ‘—ò  ›‰ ¬Ê—ê” —¬—Ì« _  ·›‰ :" & "  +982188554455,+982188554466,+982188554477,+982188554488"
                    ElseIf strDelegate = "56" Then
                        mdifrm.StatusBar1.Panels(1).Text = "In Order Of MoeinReklam Co " & " | " & " 07708615501 - 07708615502 - 07480151660 "    ' & "«Ã—« :  ‘—ò  ›‰ ¬Ê—ê” —¬—Ì« _  ·›‰ :" & "  +982188554455,+982188554466,+982188554477,+982188554488"
                    Else
                        mdifrm.StatusBar1.Panels(1).Text = " ‘⁄»Â :" & CurrentBranchName & " | " & "«Ã—«: ‘—ò  „Â‰œ”Ì ›‰ ¬Ê—ê” —¬—Ì« _  ·›‰ :" & "  +982188554455,+982188554466,+982188554477,+982188554488"
                      '  mdifrm.StatusBar1.Panels(1).Text = "«Ã—«: ‘—ò  „Â‰œ”Ì ›‰ ¬Ê—ê” —¬—Ì« " '& "  +982188554455,+982188554466,+982188554477,+982188554488" ' " ‘⁄»Â :" & CurrentBranchName & " | " &
                    End If
                    mdifrm.StatusBar1.Panels(3).Text = clsDate.shamsi(Date)
                    mdifrm.StatusBar1.Panels(4).Text = mdifrm.StatusBar1.Panels(4).Text & " = " & aa
                    mdifrm.StatusBar1.Panels(5).Text = mdifrm.StatusBar1.Panels(5).Text & " = " & bb
                    mdifrm.StatusBar1.Panels(6).Text = mdifrm.StatusBar1.Panels(6).Text & SoftwareVersion & "_" & CurrentScriptNo
                Else
                    If strDelegate = "56" Then
                        mdifrm.StatusBar1.Panels(1).Text = "In Order Of MoeinReklam Co " & " | " & " 07708615501 - 07708615502 - 07480151660 "    ' & "«Ã—« :  ‘—ò  ›‰ ¬Ê—ê” —¬—Ì« _  ·›‰ :" & "  +982188554455,+982188554466,+982188554477,+982188554488"
                    Else
                        mdifrm.StatusBar1.Panels(1).Text = " Branch :" & CurrentBranchName & " | " & "WWW.FGARYA.COM / TEl :" & "  +982188554455,+982188554466,+982188554477,+982188554488"
                    End If
                    mdifrm.StatusBar1.Panels(3).Text = clsDate.shamsi(Date)
                    mdifrm.StatusBar1.Panels(4).Text = mdifrm.StatusBar1.Panels(4).Text & " = " & aa
                    mdifrm.StatusBar1.Panels(5).Text = mdifrm.StatusBar1.Panels(5).Text & " = " & bb
                    mdifrm.StatusBar1.Panels(6).Text = mdifrm.StatusBar1.Panels(6).Text & SoftwareVersion & "_" & CurrentScriptNo
                 End If
                 If AccountYear <> left(clsDate.shamsi(Date), 4) Then
                     frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ »« ”«· „«·Ì -" & AccountYear & " - òÂ „Œ«·›  «—ÌŒ ›⁄·Ì «”  !«œ«„Â œÂÌœø "
                     frmMsg.fwlblMsg.Caption = frmMsg.fwlblMsg.Caption & "œ— ’Ê—  «‰ Œ«» »·Ì ”Ì” „ »« Â„Ì‰ ”«· „«·Ì «œ«„Â „Ì œÂœ"
                     frmMsg.fwlblMsg.Caption = frmMsg.fwlblMsg.Caption & " Ê œ— ’Ê—  «‰ Œ«» ŒÌ— ”Ì” „ ”«· „«·Ì ÃœÌœ —« ”«Œ Â Ê »« ”«· „«·Ì ÃœÌœ «œ«„Â „Ì œÂœ"
                     frmMsg.fwBtn(0).ButtonType = flwButtonOk
                     frmMsg.fwBtn(0).Caption = "»·Ì"
                     frmMsg.fwBtn(1).Visible = flwButtonCancel
                     frmMsg.fwBtn(1).Caption = "ŒÌ—"
                     frmMsg.fwBtn(1).Default = True
                     frmMsg.Show vbModal
                     If mvarMsgIdx = vbNo Then
                         AccountYear = left(clsDate.shamsi(Date), 4)
                         ReDim Parameter(2) As Parameter
                         Parameter(0) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                         Parameter(1) = GenerateInputParameter("@UserId", adSmallInt, 2, mvarCurUserNo)
                         Parameter(2) = GenerateOutputParameter("@intStatus", adInteger, 4)
                         
                         Result = RunParametricStoredProcedure("Insert_tAccountYears", Parameter)
                         
                         If Result = 1 Then
                             frmMsg.fwlblMsg.Caption = "”«· „«·Ì ÃœÌœ -" & AccountYear & "- ”«Œ Â ‘œ."
                             frmMsg.fwBtn(0).Visible = False
                             frmMsg.fwBtn(1).ButtonType = flwButtonOk
                             frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                             frmMsg.Show vbModal
                         End If
                        frmMsg.fwlblMsg.Caption = "»«Ìœ ﬂ«·«Â«Ì «‰»«— »—«Ì «Ì‰ ”«· „«·Ì «÷«›Â ‘Êœ."
                        frmMsg.fwBtn(0).Visible = False
                        frmMsg.fwBtn(1).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                        frmMsg.Show vbModal
                        frmInventory_level1.Show
                        Sleep 1000
                        frmInventory_level1.cmbBranch.ListIndex = 0
                        Sleep 1000
                        frmInventory_level1.cmbInventory.ListIndex = 0
                        Sleep 5000
                        frmInventory_level1.AddGoodstoInventory
                        Sleep 3000
                        frmInventory_level1.ExitForm
                        
                        frmMsg.fwlblMsg.Caption = " ﬂ«·«Â«Ì «‰»«—„—ﬂ“Ì »—«Ì «Ì‰ ”«· „«·Ì «÷«›Â ‘œ." & "»—«Ì ”«Ì— «‰»«—Â« „œÌ— ”Ì” „ «ﬁœ«„ ‰„«Ìœ."
                        frmMsg.fwBtn(0).Visible = False
                        frmMsg.fwBtn(1).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                        frmMsg.Show vbModal
                        Unload frmInput
                        ReDim Parameter(0) As Parameter
                        Parameter(0) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                        RunParametricStoredProcedure "Update_tAccountYears", Parameter
                        SaveSetting strMainKey, "SalMali", "SalMali", AccountYear
                  End If
                
                 End If
                
                If Rst.State <> 0 Then Rst.Close
                Dim f As New FileSystemObject
                Dim filetemp As New FileSystemObject
                Dim tempstring As TextStream
                Dim Str As String
                Dim IsFileExist As Boolean
                UserSettingFile = App.Path & "\Setting\User" & mvarCurUserNo & ".txt"
                IsFileExist = filetemp.FileExists(UserSettingFile)
                
                If IsFileExist = False Then
                   SetDefaultUserSettingFile
                   MsgBox "User Setting File Did Not Exist" & vbCrLf & "Default User Setting File Created"
                End If
                
                MojodiControlFlag = clsStation.MojodiControlDefault
                
                If clsStation.SellerCaption = "" Then
                    clsStation.SellerCaption = "›—Ê‘‰œÂ"
                End If
                
                ReDim Parameter(2) As Parameter
                Parameter(0) = GenerateInputParameter("@intUserNo", adInteger, 4, mvarCurUserNo)
                Parameter(1) = GenerateInputParameter("@intActionUserNo", adInteger, 4, 1)
                Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                RunParametricStoredProcedure "Insert_tblTotal_UserHistory", Parameter
                
                MainPriceType = clsStation.PriceType
                
                If CurrentDateNumber > LastDateNumber And Station_IsServer = True Then
                    frmDisMsg.Timer1.Interval = 1000
                    frmDisMsg.lblMessage = "”Ì” „ œ— Õ«· „— » ”«“Ì œ«œÂ Â« „Ì »«‘œ "
                    frmDisMsg.lblMessage = frmDisMsg.lblMessage & vbLf & "·ÿ›« „‰ Ÿ— »„«‰Ìœ "
                    frmDisMsg.Show
                    DoEvents
                    On Error Resume Next
                    RunNonParametricStoredProcedure "Total_Reindex"
                    On Error GoTo 0
                    frmDisMsg.Timer1.Enabled = True
                End If
                
                If CurrentDateNumber > LastDateNumber And ClsFormAccess.CashRemaining = True And clsArya.ExternalAccounting = False Then
                    frmCashSharzh.Show vbModal
                    Unload frmCashSharzh
                End If
                If CurrentDateNumber > LastDateNumber And Station_IsServer = True And clsStation.AutoBackup = False Then
                   ShowMessage "¬Ì« „«Ì· »Â ê—› ‰ ‰”ŒÂ Å‘ Ì»«‰ «“ «ÿ·«⁄«  ŒÊœ Â” Ìœø", True, True, "»·Ì", "ŒÌ—"
                   If mvarMsgIdx = vbYes Then
                        frmDisMsg.Timer1.Interval = 1000
                        frmDisMsg.lblMessage = "”Ì” „ œ— Õ«· ê—› ‰ ‰”ŒÂ Å‘ Ì»«‰ «“ œ«œÂ Â« „Ì »«‘œ "
                        frmDisMsg.lblMessage = frmDisMsg.lblMessage & vbLf & "·ÿ›« „‰ Ÿ— »„«‰Ìœ "
                        frmDisMsg.Show
                        DoEvents
                        On Error Resume Next
                        mdifrm.AutoBackup
                        On Error GoTo 0
                   End If
                End If
                
                If clsArya.ExternalAccounting = True Or HasMiniAcc = True Then
                    Accounting.SendVariables strMainKey, clsArya.DBLogin, SqlPass, clsArya.DbName, clsArya.ServerName, CStr(AccountYear), Trim(clsArya.Company), CurrentBranch, mvarCurUserNo, mVarAccessLevel, mvarTafsili, AryaSettingFile, "lemon74300347nomel", App.Path, clsArya.StationNo, clsStation.PosModel, HasPcPos
    ''                Dim Form As New FGArya_Samar_Accounting.FormShowHandler
    ''                Form.SetSaleSale AccountYear, CurrentBranch, mvarCurUserNo, DataBaseName, clsArya.DBLogin, clsArya.ServerName, Trim(clsArya.Company)
                    If CurrentDateNumber > LastDateNumber And Station_IsServer = True And (LCase(clsArya.ServerName) = LCase(MachineName) Or clsArya.ServerName = ".") Then
                        ShowMessage "¬Ì« „«Ì· »Â ê—› ‰ ê“«—‘ çﬂÂ«Ì œ—Ì«› ‰Ì  «  «—ÌŒ ›—œ« Â” Ìœø", True, True, "»·Ì", "ŒÌ—"
                        If mvarMsgIdx = vbYes Then
                            Accounting.ChequePrintingDll
                        End If
                    End If
                    If CurrentDateNumber > LastDateNumber And Station_IsServer = True And (LCase(clsArya.ServerName) = LCase(MachineName) Or clsArya.ServerName = ".") Then
                        ShowMessage "¬Ì« „«Ì· »Â ê—› ‰ ê“«—‘ çﬂÂ«Ì Å—œ«Œ ‰Ì  «  «—ÌŒ ›—œ« Â” Ìœø", True, True, "»·Ì", "ŒÌ—"
                        If mvarMsgIdx = vbYes Then
                            Accounting.ChequePaymentPrintingDll
                        End If
                    End If
                    Dim TafsiliFlag As Boolean
                    If mvarTafsili = 0 Then
                        ShowMessage " ò«—»— ›⁄·Ì »—«Ì À»  œ— ”Ì” „ Õ”«»œ«—Ì œ«—«Ì òœ  ›’Ì·Ì ‰Ì”  . Ìò»«— ò«—»— „Ê—œ ‰Ÿ— —« ÊÌ—«Ì‘ ò‰Ìœ Ì« Â„Â ò«—»—«‰ —« »Â ”Ì” „ Õ”«»œ«—Ì «÷«›Â ò‰Ìœ ", True, False, "ﬁ»Ê·", ""
                        frmPer.Show
                        TafsiliFlag = True
                    End If
                End If
                If TafsiliFlag = False Then
                    If clsStation.StartUpFormDefault = 0 Then
                        If ClsFormAccess.frmInvoice = True Then
                            frmInvoice.Show
                            VarActForm = "frmInvoice"
                            frmInvoice.Add      'Because frmCashSharzh added and no go to Activate
                        End If
                    ElseIf clsStation.StartUpFormDefault = 1 Then
                        frmPurchase.Show
                    End If
                End If
            End If
         End If
    End If
    If Rst.State = adStateOpen Then Rst.Close: Set Rst = Nothing

    Call PresetScreenSaver
    mdifrm.tmrScreenSaver.Enabled = True

    Dim nodX
    For Each nodX In frmGroupMenu.trMenu.Nodes
        nodX.Expanded = False
       ' nodX.EnsureVisible
    Next nodX

     If HasAryaSms = True And clsStation.AryaSmsPanel = True Then
         Dim AryaSms
         Set AryaSms = CreateObject("AryaSms.clsMonitoring")
         AryaSms.SendVariables strMainKey, clsArya.ServerName, clsArya.DbName, clsArya.DBLogin, clsArya.DBPass, Trim(clsArya.Company), "lemon74300347", App.Path
         AryaSms.ShowForms
     End If


Exit Sub

ErrorHandler:
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    If Rst.State = adStateOpen Then Rst.Close: Set Rst = Nothing
    Set clsDate = Nothing
End Sub
        
Private Sub Form_Activate()
     '  Normal & Restaurant
        Me.BackColor = &HFFC0C0           ' &H9000&
'        Me.lblAddress.ForeColor = &H40&
'        Me.lblAddress.BackColor = &HFFC0C0
        Label1.ForeColor = vbBlue
        Label2.ForeColor = vbBlue
        Label3.ForeColor = vbBlue
        Label4.ForeColor = vbBlue
    '    Me.Frame1.BackColor = &H9000&
        Me.Frame2.BackColor = &HFFC0C0
        Me.Frame3.BackColor = &HFFC0C0
        Me.cmbBranch.BackColor = &HFFC0C0
        Me.cmbSalMali.BackColor = &HFFC0C0

'        Me.LblSoftwareRegistrationNotice.BackColor = &HFFC0C0
'        Me.LblSoftwareRegistrationNotice.ForeColor = vbRed

    On Error Resume Next
    If TxtUserName.Text <> "" Then
        txtPassword.SetFocus
    Else
        TxtUserName.SetFocus
    End If
    SetKbLayout LANG_EN_US
    On Error GoTo 0
    
    If Trim(clsArya.Company) = "" Then End
    If clsArya.HardLock = False Then
        If clsArya.LimitedVersion = True And HardLockFlagTrial = False Then
'            lblAddress.Caption = "       ‰—„ «›“«—Õ«÷—  Õ  ﬁÊ«‰Ì‰ ‘Ê—«Ì⁄«·Ì «‰›Ê—„« Ìò ò‘Ê— „Ì »«‘œ" & "«Ì‰ ‰”ŒÂ «“ »—‰«„Â »—«Ì «” ›«œÂ " & Trim(clsArya.Company) & "   ÂÌÂ ê—œÌœÂ Ê ﬁ«»· òÅÌ Ê  òÀÌ—  „Ì »«‘œ. "
'            lblAddress.Caption = lblAddress.Caption & " «” ›«œÂ «“ «Ì‰ ‰”ŒÂ œ«—«Ì „ÕœÊœÌ  „Ì »«‘œ Ê Å” «“ „‘«ÂœÂ ﬁ«»·Ì  Â«Ì Ê—é‰ Â«Ì „Œ ·› ‰—„ «›“«—»—«Ì  ÂÌÂ ‰”ŒÂ «’·Ì »« ‰„«Ì‰œê«‰ ›—Ê‘ ‘—ò   „«” Õ«’· ›—„«∆Ìœ"
            strPolicy = "       ‰—„ «›“«—Õ«÷—  Õ  ﬁÊ«‰Ì‰ ‘Ê—«Ì⁄«·Ì «‰›Ê—„« Ìò ò‘Ê— „Ì »«‘œ" & "«Ì‰ ‰”ŒÂ «“ »—‰«„Â »—«Ì «” ›«œÂ " & Trim(clsArya.Company) & "   ÂÌÂ ê—œÌœÂ Ê ﬁ«»· òÅÌ Ê  òÀÌ—  „Ì »«‘œ. "
            strPolicy = strPolicy & " «” ›«œÂ «“ «Ì‰ ‰”ŒÂ œ«—«Ì „ÕœÊœÌ  „Ì »«‘œ Ê Å” «“ „‘«ÂœÂ ﬁ«»·Ì  Â«Ì Ê—é‰ Â«Ì „Œ ·› ‰—„ «›“«—»—«Ì  ÂÌÂ ‰”ŒÂ «’·Ì »« ‰„«Ì‰œê«‰ ›—Ê‘ ‘—ò   „«” Õ«’· ›—„«∆Ìœ"
        ElseIf clsArya.LimitedVersion = True And HardLockFlagTrial = True Then
            strPolicy = "       ‰—„ «›“«—Õ«÷—  Õ  ﬁÊ«‰Ì‰ ‘Ê—«Ì⁄«·Ì «‰›Ê—„« Ìò ò‘Ê— „Ì »«‘œ" & "«Ì‰ ‰”ŒÂ «“ »—‰«„Â »—«Ì «” ›«œÂ œ— " & Trim(clsArya.Company) & " Ê«ﬁ⁄ œ— " & Trim(clsArya.CustomerAddres) & "   ÂÌÂ ê—œÌœÂ Ê ﬁ«»· òÅÌ Ê  òÀÌ—  „Ì »«‘œ. "
        Else
            strPolicy = "       ‰—„ «›“«—Õ«÷—  Õ  ﬁÊ«‰Ì‰ ‘Ê—«Ì⁄«·Ì «‰›Ê—„« Ìò ò‘Ê— „Ì »«‘œ" & "«Ì‰ ‰”ŒÂ «“ »—‰«„Â »—«Ì «” ›«œÂ œ— " & Trim(clsArya.Company) & " Ê«ﬁ⁄ œ— " & Trim(clsArya.CustomerAddres) & "   ÂÌÂ ê—œÌœÂ  " & "Ê Â—êÊ‰Â òÅÌ »—œ«—Ì Ê«” ›«œÂ €Ì— ﬁ«‰Ê‰Ì «“ ¬‰  ﬁ«»· ÅÌê—œ „Ì »«‘œ "
        End If
    Else
        strPolicy = "       ‰—„ «›“«—Õ«÷—  Õ  ﬁÊ«‰Ì‰ ‘Ê—«Ì⁄«·Ì «‰›Ê—„« Ìò ò‘Ê— „Ì »«‘œ" & "«Ì‰ ‰”ŒÂ «“ »—‰«„Â »—«Ì «” ›«œÂ œ— " & Trim(clsArya.Company) & " Ê«ﬁ⁄ œ— " & Trim(clsArya.CustomerAddres) & "   ÂÌÂ ê—œÌœÂ  " & "Ê Â—êÊ‰Â òÅÌ »—œ«—Ì Ê«” ›«œÂ €Ì— ﬁ«‰Ê‰Ì «“ ¬‰  ﬁ«»· ÅÌê—œ „Ì »«‘œ "
    End If
    
    lblAddress = strPolicy

    lblSpec.Caption = Format(clsArya.CustomerId, "00000") & "-" & strCategory & "-" & strDelegate
    Select Case intVersion
        Case 3
            LblVersion.Caption = "„Ì‰Ì"
        Case 0
            LblVersion.Caption = "„ Ê”ÿ"
        Case 1
            LblVersion.Caption = "ÅÌ‘—› Â"
        Case 2
            LblVersion.Caption = "ÊÌéÂ"
        Case 4
            LblVersion.Caption = "«·„«”"
    End Select
    
    LblSerialNo.Caption = clsArya.HardLockSerialNo
    
    ShamsiDateName = clsDate.shamsi(Date)
'    ShamsiDateName = FarDate1.MiladiToShamsi(Date)
'    ShamsiDateName = Left(ShamsiDateName, 4) & "/" & Mid(ShamsiDateName, 5, 2) & "/" & Right(ShamsiDateName, 2)
    If clsArya.MiladiDate = 0 Then CurrentDateNumber = DateToNumber8(Mid(ShamsiDateName, 3))
    On Error Resume Next
    LastDateNumber = Val(GetSetting(strMainKey, "LastDateNumber", "LastDateNumber"))
    On Error GoTo 0
    SaveSetting strMainKey, "LastDateNumber", "LastDateNumber", CurrentDateNumber
    
    
    FWLblDate.Caption = clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)) & " " & ShamsiDateName
    
    ReDim Parameters(0) As Parameter
    Parameters(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    RunParametricStoredProcedure "NonCustomerCheck", Parameters
     
    ReDim Parameters(0) As Parameter
    Parameters(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    RunParametricStoredProcedure "NonSupplierCheck", Parameters
             
    Dim Result As Integer
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(1) = GenerateOutputParameter("@Count", adInteger, 4)
    
    Result = RunParametricStoredProcedure("DefaultUserCheck", Parameter)
    If Result = 0 Then
        Unload frmMsg
        frmPer.Show
    Else
        TempPerFlag = True
    End If
    
    AutoHavale = 1

End Sub
Private Sub FillVersion()
    Dim i As Integer
    CmbVersion.Clear
    CmbVersion.AddItem "”«œÂ"
    CmbVersion.ItemData(0) = EnumVersion.Min
    CmbVersion.AddItem "„ Ê”ÿ"
    CmbVersion.ItemData(1) = EnumVersion.Normal
    CmbVersion.AddItem "ÅÌ‘—› Â"
    CmbVersion.ItemData(2) = EnumVersion.Silver
    CmbVersion.AddItem "ÊÌéÂ"
    CmbVersion.ItemData(3) = EnumVersion.gold
    CmbVersion.AddItem "«·„«”"
    CmbVersion.ItemData(4) = EnumVersion.Diamond
     
    If clsArya.LimitedVersion = True Or DebugMode = True Then
        CmbVersion.ListIndex = 4
        CmbVersion.Enabled = True
    Else
        For i = 0 To CmbVersion.ListCount
            If CmbVersion.ItemData(i) = intVersion Then
               CmbVersion.ListIndex = i
               Exit For
            End If
        Next i
        CmbVersion.Enabled = False
    End If
    
End Sub

Private Sub TxtDataSource_Change()
    If formloadFlag = False Then Exit Sub
    MakeConnectionString

End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        Command1_Click 0
    End If
End Sub

''''Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''''    If KeyCode = 8 And Shift = 0 Then
''''        BtnKeypad_Click (10)
''''    Else
''''        Command1_Click (0)
''''    End If
''''End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
Dim ii As Integer
    
    VarActForm = Me.Name
    
    Dim hMenu As Long
    
    hMenu = GetSystemMenu(Me.hwnd, False)
    
    DeleteMenu hMenu, 6, MF_BYPOSITION
    
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

        
''############
''   «» œ« ¬—Ì« ” Ì‰ê œ— œ—«ÌÊ ”Ì çò „Ì ‘ÊœÊ «ê— ÅÌœ« ‰‘œ œ— „”Ì— Ã«—Ì œ‰»«· „Ì ê—œœ
'    AryaSettingFile = "C:\Aryasetting.txt"
'    IsFileExist = filetemp.FileExists(AryaSettingFile)
'
'    If IsFileExist = False Then
'        AryaSettingFile = App.Path & "\Aryasetting.txt"
'        IsFileExist = filetemp.FileExists(AryaSettingFile)
'        If IsFileExist = False Then
'            setDefaultAryaSettingFile
'            If clsArya.LimitedVersion = False Then MsgBox "Arya Setting File " & vbCrLf & "œ— „”Ì—  " & App.Path & "\Aryasetting.txt" & "«ÌÃ«œ ê—œÌœ"
'        End If
'    End If

'############ »—«Ì «Ì‰òÂ «„ò«‰ «Ã—«Ì œÊ »—‰«„Â Â„ “„«‰ œ— ”Ì” „ ÊÃÊœ œ«‘ Â »«‘œ
'   «» œ« ¬—Ì« ” Ì‰ê œ— „”Ì— Ã«—Ì çò „Ì ‘ÊœÊ «ê— ÅÌœ« ‰‘œ œ— œ—«ÌÊ ”Ì œ‰»«· „Ì ê—œœ
    AryaSettingFile = App.Path & "\Aryasetting.txt"
    IsFileExist = filetemp.FileExists(AryaSettingFile)
    
    If IsFileExist = False Then
        AryaSettingFile = "C:\Aryasetting.txt"
        IsFileExist = filetemp.FileExists(AryaSettingFile)
        If IsFileExist = False Then
            setDefaultAryaSettingFile
            If clsArya.LimitedVersion = False Then MsgBox "Arya Setting File " & vbCrLf & "œ— „”Ì—  " & "C:\Aryasetting.txt" & "«ÌÃ«œ ê—œÌœ"
        End If
    End If
'############
    
    If clsArya.HardLockSerialNo = "93061701000" Then   ' Palladium Mall Server
        AryaSettingFile = App.Path & "\Aryasetting.txt"
        IsFileExist = filetemp.FileExists(AryaSettingFile)
        If IsFileExist = False Then
            MsgBox "Arya Setting File " & vbCrLf & "œ— „”Ì—  " & App.Path & "\Aryasetting.txt" & "ÊÃÊœ ‰œ«—œ Ê ”Ì” „ ﬁ«œ— »Â «œ«„Â ‰Ì” "
            End
        Else
            MsgBox "Arya Setting File " & vbCrLf & "œ— „”Ì—  " & App.Path & "\Aryasetting.txt" & "‘‰«”«ÌÌ ‘œ"
            clsArya.Class_Initialize
            
        End If
    End If
    CenterCenterinSecondScreen Me
    
    If filetemp.FolderExists(App.Path & "\Setting") = False Then
        filetemp.CreateFolder (App.Path & "\Setting")
    End If
    
    StationSettingFile = App.Path & "\Setting\Station" & clsArya.StationNo & ".txt"
    IsFileExist = filetemp.FileExists(StationSettingFile)
    
    If IsFileExist = False Then
      SetDefaultStationSettingFile
      MsgBox "Station Setting File Did Not Exist" & vbCrLf & "Default Station Setting File Created"
    
    End If
    
    InvoiceSettingFile = App.Path & "\Setting\Invoice" & ".txt"
    IsFileExist = filetemp.FileExists(InvoiceSettingFile)
    
    If IsFileExist = False Then
      SetDefaultInvoiceSettingFile
      MsgBox "Invoice Setting File Did Not Exist" & vbCrLf & "Default Invoice Setting File Created"
    
    End If
    
    GoodMenuSettingFile = App.Path & "\Setting\GoodMenu" & ".txt"
    IsFileExist = filetemp.FileExists(GoodMenuSettingFile)
    
    If IsFileExist = False Then
      SetDefaultGoodMenuSettingFile
      MsgBox "GoodMenu Setting File Did Not Exist" & vbCrLf & "Default GoodMenu Setting File Created"
    
    End If
    
'    AccountingSettingFile = App.Path & "\Setting\Accounting.txt"
'    If Not filetemp.FileExists(AccountingSettingFile) Then
'        setDefaultAccountingSettingFile
'        MsgBox "Accounting file didn't exist" & vbCrLf & "Default accounting setting file created"
'    End If
    
'    If Trim(clsArya.Company) = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" Then MsgBox "›«Ì· «Ã—«∆Ì Œ«„ «” ": End
    If Trim(clsArya.HardLockSerialNo) = "xxxxxxxxxxxxxxxxxx" Then MsgBox "›«Ì· «Ã—«∆Ì Œ«„ «” ": End
    
'    If filetemp.FileExists("L:\Data_Sql\Total_Cool_Data.MDF") = False Then
    If clsArya.ExternalDataBase = True Then
       If clsArya.ExternalDbPath = "" Then
          clsArya.ExternalDbPath = "L:\Data\Total_Ext.mdf"
          setAryaSettingFile
       End If
       If filetemp.FileExists(clsArya.ExternalDbPath) = False Then
           MsgBox "  ›«Ì· œ«œÂ Â« ÅÌœ« ‰‘œ "
           End
       Else
            If clsArya.ExternalDataName = "" Then
               clsArya.ExternalDataName = "Total_Ext"
               setAryaSettingFile
            End If
            clsArya.DbName = clsArya.ExternalDataName
       End If
     End If
    
    'CurrentVersion = "2.3.1"   '84/04/07
    'CurrentVersion = "3.3.1"   '84/04/19
    'CurrentVersion = "4.4.1"   '84/07/15
    'CurrentVersion = "4.4.2"   '84/07/21
    'CurrentVersion = "4.4.3"   '84/07/22
    'CurrentVersion = "4.4"   '84/08/13
    'CurrentScriptNo = 6   '      First Update :84/07/27
    'CurrentScriptNo = 8   '      First Update :84/08/07
    'CurrentScriptNo = 9   '      First Update :84/08/13
    'CurrentScriptNo = 10   '      First Update :84/08/21
    'CurrentScriptNo = 11   '      First Update :84/09/05
    'CurrentScriptNo = 12   '      First Update :84/09/19
    'CurrentScriptNo = 13   '      First Update :84/09/23
    'CurrentScriptNo = 14   '      First Update :84/10/03
    'CurrentScriptNo = 15   '      First Update :84/10/09
    'CurrentScriptNo = 16   '      First Update :84/10/24
    'CurrentScriptNo = 17   '      First Update :84/11/14
    'CurrentScriptNo = 18   '      First Update :84/11/29
    'CurrentScriptNo = 19   '      First Update :84/12/12
    'CurrentScriptNo = 20   '      First Update :84/12/19
    'CurrentScriptNo = 22   '      First Update :85/02/19
    'CurrentScriptNo = 23   '      First Update :85/02/28
    
    'DatabaseVersion = 4   '85/10/29
    'SoftwareVersion = 25   '85/10/29
    'SoftwareVersion = 26   '85/12/25
    SoftwareVersion = 26   '85/12/25
    'CurrentScriptNo = 6   '      First Update :85/11/10
    'CurrentScriptNo = 7   '      First Update :85/12/24
    'CurrentScriptNo = 1   '      First Update :85/12/25
    'CurrentScriptNo = 2   '      First Update :86/01/03
    'CurrentScriptNo = 3   '      First Update :86/01/08
    'CurrentScriptNo = 4   '      First Update :86/01/15
    'CurrentScriptNo = 5   '      First Update :86/01/27
    'CurrentScriptNo = 6   '      First Update :86/07/10
    'CurrentScriptNo = 7   '      First Update :86/09/26
    'CurrentScriptNo = 8   '      First Update :87/06/27
    'CurrentScriptNo = 9   '      First Update :88/11/08
    'CurrentScriptNo = 10   '      First Update :89/02/10
    'CurrentScriptNo = 11   '      First Update :89/09/18
    'CurrentScriptNo = 12   '      First Update :89/11/01
    'CurrentScriptNo = 13   '      First Update :90/05/01 For Sasad
    'CurrentScriptNo = 14   '      First Update :90/08/10
   ' CurrentScriptNo = 15   '      First Update :90/12/10
    CurrentScriptNo = 16   '      First Update :92/10/15
'
    If clsArya.DBLogin = "" Then
        clsArya.DBLogin = "sa"
    End If
    
    If clsArya.DbName <> "" Then
        strConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & clsArya.DBLogin & "; Password = " & SqlPass & " ; Initial Catalog = " & clsArya.DbName & ";Data Source=" & clsArya.ServerName
      '  strConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & clsArya.DBLogin & "; Password = " & SqlPass & " ; Initial Catalog = Master" & ";Data Source=" & clsArya.ServerName
    Else
        strConnectionString = ""
    End If
    SystemFolderName = GetSystemPath64 '
    If SystemFolderName = "" Then SystemFolderName = GetSystemPath     ' FWSysInfo1.System
    MachineLocalIp = Winsock1.LocalIP
    
    TxtUserName.Text = GetSetting(strMainKey, "TxtUserName", "TxtUserName")
    If clsArya.BranchView = False Then
        Label1.Visible = False
        Label4.Visible = False
        cmbBranch.Visible = False
        cmbSalMali.Visible = False
    End If
    
    TxtDataSource.Text = clsArya.ServerName
    FillDataBase
    MakeConnectionString
    FillCmbStation
    FillBranch
    FillSalMali
    FillVersion
    
    If DebugMode = False Then SetKbLayout LANG_Pr_IR
    '####################
    ''Check in command_click
''''    If PosConnection.State = 0 Then PosConnection.Open strConnectionString
''''    If rctmp.State = adStateOpen Then rctmp.Close
''''    rctmp.Open "Select * from  dbo.tStations Where (StationType &  32  = 32 ) And StationID = " & clsArya.StationNo, PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
''''    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
''''           Station_IsAccounting = True
''''    Else
''''       Station_IsAccounting = False
''''    End If
    
    '''#############
 
    Set Tiny1 = CreateObject("Tiny.TINYCtrl.1")
    Set TinyEvent = Tiny1       ' necwssary for fire event
    
    If clsArya.LimitedVersion = True Then
''''·«“„ ‰Ì”  ﬁ›· —« çò ò‰œ
        HardLockFlagTrial = False
        Station_IsAccounting = True
'        Call HardLockCheckTrial
'        If HardLockFlagTrial = True Then
'            LblLimited.Visible = False
'            LblSerialNo.Caption = clsArya.HardLockSerialNo  ' New Serial from Lock
'            CmbVersion.Enabled = False
'            For i = 0 To CmbVersion.ListCount
'                If CmbVersion.ItemData(i) = intVersion Then
'                   CmbVersion.ListIndex = i
'                   Exit For
'                End If
'            Next i
'            SetTooltipText
'        Else
            If Trim(clsArya.Company) = "" Then clsArya.Company = "¬—Ì« "
            If Trim(clsArya.CustomerAddres) = "" Then clsArya.CustomerAddres = " Â—«‰"
            SetTooltipText_Demo
            cmbDataBase.Enabled = False
            LblLimited.Visible = True
            CmdSpec.Visible = False
            CmbVersion.Enabled = True
            IsHelp = True
            'Load frmBrowser_Demo
            frmBrowser_Demo.Show
            IsHelp = False
'        End If
    Else
        SetTooltipText
    End If
   
    If Trim(clsArya.Company) = "" Then
        frmInput.fwlblInput.Caption = "·ÿ›« ‰«„ ŒÊœ —« Ê«—œ ‰„«ÌÌœ "
        frmInput.btnCancel.Visible = False
        frmInput.Show vbModal
        If Trim(mvarInput) <> "" Then
            Set tempstring = f.OpenTextFile(AryaSettingFile, ForAppending, False, TristateFalse)
            clsArya.Company = mvarInput
            Dim CustName As String
            CustName = "CustomerName =" & clsArya.Company
            tempstring.WriteLine (CustName)
            tempstring.Close
        End If
    End If
    
    If Trim(clsArya.CustomerAddres) = "" Then
        frmInput.fwlblInput.Caption = "·ÿ›« ¬œ—” ŒÊœ —« Ê«—œ ‰„«ÌÌœ "
        frmInput.btnCancel.Visible = False
        frmInput.Show vbModal
    
        If Trim(mvarInput) <> "" Then
            Set tempstring = f.OpenTextFile(AryaSettingFile, ForAppending, False, TristateFalse)
            clsArya.CustomerAddres = mvarInput
            Dim CustAddress As String
            CustAddress = "CustomerAddress =" & clsArya.CustomerAddres
            tempstring.WriteLine (CustAddress)
            tempstring.Close
        End If
    End If
    
    frmInput.btnCancel.Visible = True
    
    formloadFlag = True
    If filetemp.FileExists(SystemFolderName & "\osk.exe") = True Then
'        Shell SystemFolderName & "\osk.exe", vbNormalFocus
        Dim lHwnd As Long, lpClassName As String, retval As Long
'        Shell SystemFolderName & "\osk.exe", vbNormalFocus
        lHwnd = FindWindow(vbNullString, "On-Screen Keyboard")
        lpClassName = Space(256)
        retval = GetClassName(lHwnd, lpClassName, 256)
        SetWindowPos lHwnd, 0, 10, 10, 0, 0, SWP_NOZORDER Or SWP_SHOWWINDOW Or SWP_NOSIZE      '    ShellExecute Me.hwnd, vbNullString, SystemFolderName & "\osk.exe", vbNullString, "C:\", 2
    End If
'    frmKeyBoard.Show
'    frmKeyBoard.Top = Me.Top + Me.Height + 100
    
    If strDelegate = "56" Then cmd_FgaryaSite.Caption = "www.MoeinReklam.com": Image_Fgarya.Visible = False: Image_Moein.Visible = True
    
    Exit Sub
ErrorHandler:
    MsgBox err.Description
    Select Case err.Number
        Case -2147467259
            End
        Case Else
            Resume Next
    End Select
End Sub
Private Sub FillDataBase()
    Dim i As Integer
    i = 0
    cmbDataBase.Clear
    Set rctmp = RunStoredProcedure2RecordSet("sp_MShasdbaccess")
'    If PosConnection.State = 0 Then PosConnection.Open strConnectionString
'    rctmp.Open "select name as DbName From master.dbo.sysdatabases Where has_dbaccess(Name) = 1 order by name", PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
' '   rctmp.Open "SELECT  DISTINCT CATALOG_NAME as DbName FROM  INFORMATION_SCHEMA.SCHEMATA WHERE   CATALOG_NAME NOT IN ('master', 'tempdb', 'model', 'msdb') order by CATALOG_NAME", PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While rctmp.EOF = False
        If rctmp!DbName <> "master" And rctmp!DbName <> "tempdb" And rctmp!DbName <> "model" And rctmp!DbName <> "msdb" And rctmp!DbName <> "Pubs" Then
            cmbDataBase.AddItem rctmp!DbName
            cmbDataBase.ItemData(cmbDataBase.NewIndex) = i
            i = i + 1
        End If
        rctmp.MoveNext
    Loop
    If rctmp.State = adStateOpen Then If rctmp.State = adStateOpen Then rctmp.Close
    For i = 0 To cmbDataBase.ListCount - 1
        cmbDataBase.ListIndex = i
        If LCase(Trim(clsArya.DbName)) = LCase(cmbDataBase.Text) Then
            Exit For
        End If
    Next
    
End Sub
Private Sub cmbDataBase_Click()
    If formloadFlag = False Then Exit Sub
    MakeConnectionString
    FillCmbStation
    FillBranch
    FillSalMali
End Sub
Private Sub MakeConnectionString()
    
    strConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & clsArya.DBLogin & "; Password = " & SqlPass & " ; Initial Catalog = " & cmbDataBase.Text & ";Data Source=" & clsArya.ServerName & "; Connect Timeout=300 "
    AccstrConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & clsArya.DBLogin & "; Password = " & SqlPass & " ; Initial Catalog = " & cmbDataBase.Text & ";Data Source=" & clsArya.ServerName & "; Connect Timeout=300 "
    CrystallConnection = "Data Source=" & clsArya.ServerName & ";UID=" & clsArya.DBLogin & ";PWD=" & SqlPass & ";DSQ= " & cmbDataBase.Text & ";"
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    clsArya.DbName = cmbDataBase.Text
End Sub

Private Sub FillSalMali()
    
    AccountYear = Val(GetSetting(strMainKey, "SalMali", "SalMali"))
    If AccountYear = 0 Then
        AccountYear = left(clsDate.shamsi(Date), 4)  'Left(FarDate1.MiladiToShamsi(Date), 4)
        SaveSetting strMainKey, "SalMali", "SalMali", AccountYear
    End If
    cmbSalMali.Clear
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_tAccountYears")
    Do While rctmp.EOF = False
        cmbSalMali.AddItem rctmp!AccountYear
        rctmp.MoveNext
    Loop
    For i = 0 To cmbSalMali.ListCount - 1
        cmbSalMali.ListIndex = i
        If AccountYear = cmbSalMali.Text Then
            Exit For
        End If
    Next
    'If cmbSalMali.ListCount > 0 Then cmbSalMali.ListIndex = 0
    rctmp.Close
End Sub
Private Sub FillBranch()
    
    cmbBranch.Clear
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_Branches")
    Do While rctmp.EOF = False
        cmbBranch.AddItem rctmp!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = rctmp!Branch
        rctmp.MoveNext
    Loop
    rctmp.Close
''''    If cmbBranch.ListCount > 1 And clsArya.BranchView = True Then
''''        cmbBranch.Visible = True
''''        Label4.Visible = True
''''    End If
    ReDim Parameters(0) As Parameter
    Parameters(0) = GenerateInputParameter("@intStationNo", adInteger, 4, clsArya.StationNo)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_StationIP_ByStationNo", Parameters)
    If rctmp.EOF = False Then
        CurrentBranch = rctmp!Branch
    End If
    For i = 0 To cmbBranch.ListCount - 1
        If CurrentBranch = cmbBranch.ItemData(i) Then
            cmbBranch.ListIndex = i
            Exit For
        End If
    Next
    CurrentBranchName = cmbBranch.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    VarActForm = ""
    
    SaveSetting strMainKey, "TxtUserName", "TxtUserName", TxtUserName.Text
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top

   Set clsDate = Nothing
   CloseWindow "On-Screen Keyboard"
   
   'Unload frmKeyBoard
End Sub

Private Sub txtPassword_GotFocus()
    txtUserIsFocus = False
    Set objName = txtPassword
End Sub

Private Sub txtPassword_Validate(Cancel As Boolean)
        txtPassword.TabIndex = 1
        TxtUserName.TabIndex = 0
End Sub

Private Sub txtUserName_GotFocus()
    txtUserIsFocus = True
    Set objName = TxtUserName
End Sub

Private Sub TxtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        KeyCode = 0
        txtPassword.SetFocus
    End If

End Sub
