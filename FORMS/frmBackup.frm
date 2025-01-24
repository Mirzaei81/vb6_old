VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmBackup 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   Icon            =   "frmBackup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   5640
   Begin VB.Frame Frame3 
      Caption         =   " «—ÌŒ Å‘ Ì»«‰"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   5415
      Begin VB.TextBox txtDate 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   5115
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "„”Ì—›«Ì· Å‘ Ì»«‰"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   5415
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2520
         Left            =   135
         TabIndex        =   1
         Top             =   900
         Width           =   5160
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         TabIndex        =   0
         Top             =   420
         Width           =   5160
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "›«Ì· Å‘ Ì»«‰"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   5160
      Width           =   5415
      Begin VB.TextBox txtFile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   5115
      End
   End
   Begin VB.PictureBox FWLabel1 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0070570E&
      Height          =   465
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   5925
      TabIndex        =   7
      Top             =   0
      Width           =   5985
   End
   Begin FLWCtrls.FWButton fwBtn 
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   8
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   " «ÌÌœ"
      FontName        =   "B Homa"
      FontBold        =   -1  'True
      FontSize        =   9.75
      Alignment       =   1
   End
   Begin FLWCtrls.FWButton fwBtn 
      Cancel          =   -1  'True
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ButtonType      =   1
      Caption         =   "«‰’—«›"
      FontName        =   "B Homa"
      FontBold        =   -1  'True
      FontSize        =   9.75
      Alignment       =   1
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmBackup.frx":A4C2
      TabIndex        =   10
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Filename As String
Private iHeight As Integer
Private iWidth As Integer
Private clsDate As New clsDate
Private CurrentYear, Currentmonth, Monthdbname As String
Private Conn As New ADODB.Connection
Private cmd As New ADODB.command
Dim tem_str As Date
Dim dt As New clsDate

Private Sub Drive1_Change()
On Error GoTo ErrorHandler
Dir1.Path = Drive1.Drive
Exit Sub
ErrorHandler:
   ' load frmMsg
    frmMsg.fwlblMsg.Caption = "œ—«ÌÊ „Ê—œ ‰Ÿ— œ— œ” —” ‰Ì”  "
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = " «∆Ìœ"
    frmMsg.Show vbModal
    Drive1.Drive = "C:"
End Sub

Private Sub Form_Activate()
    
    VarActForm = Me.Name

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
    Dim shamsi As String
    CenterTop Me
    
    If ClsFormAccess.frmBackup = False Then
        Unload Me
    End If
    
    Conn.ConnectionString = strConnectionString
    If Conn.State <> adStateOpen Then Conn.Open
    
    Drive1.Drive = "c:\"
    Dir1.Path = "c:\"
    shamsi = clsDate.shamsi(Date)
    txtFile.Text = Trim(Mid(shamsi, 3, 2)) & Trim(Mid(shamsi, 6, 2)) & Trim(Mid(shamsi, 9, 2)) & ".fbk"
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

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set clsDate = Nothing
    If Conn.State = adStateOpen Then Conn.Close: Set Conn = Nothing
    Set mdifrm.FileCls = Nothing
    mdifrm.Toolbar1.Buttons(20).Enabled = False
    mdifrm.Toolbar1.Buttons(21).Enabled = False
'    mdifrm.Toolbar1.Buttons(27).Enabled = False
    VarActForm = ""
    modgl.mvarDeleteMsg = ""
    Unload frmBackup
    'mdifrm.Arrange 0
    mdifrm.Toolbar1.Buttons(20).Enabled = False
    mdifrm.Toolbar1.Buttons(21).Enabled = False
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    'mdifrm.PicKeyBoard.Visible = False

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


End Sub



Private Sub fwBtn_Click(index As Integer)
Dim i As Integer
Dim CurrentBackupNo As Integer
'On Error Resume Next
frmBackup.Enabled = False

If index = 0 Then
Dim f As New FileSystemObject
Dim cnt As Integer
Dim IsFileExist As Boolean
cnt = 1
    
    Select Case txtDate
        Case ""
           
          Do
                IsFileExist = f.FileExists(Dir1.Path & "\" & Mid(txtFile.Text, 1, 6) & "_" & cnt & ".fbk")
                If IsFileExist = False Then
                    Exit Do
                End If
                cnt = cnt + 1
           Loop
            With cmd
                .ActiveConnection = Conn
                .CommandType = adCmdText
                
                .CommandText = "USE master" & _
                               " EXEC sp_addumpdevice 'disk', 'tmpTotal', '" & Dir1.Path & "\" & Mid(txtFile.Text, 1, 6) & "_" & CStr(cnt) & ".fbk" & " ' " & _
                               " BACKUP DATABASE [" & clsArya.DbName & "] To tmpTotal " & _
                               " exec sp_dropdevice 'tmpTotal' "
            End With
            cmd.Execute
            cmd.Cancel
        Case Else
          Do
                IsFileExist = f.FileExists(Dir1.Path & "\" & Mid(txtFile.Text, 1, 6) & "_" & cnt & ".dbk")
                If IsFileExist = False Then
                    Exit Do
                End If
                cnt = cnt + 1
           Loop
            Me.InsertDb "tCust", "Pubs", clsArya.DbName, "Where [Date] = '" & txtDate & "'"
            Me.InsertDb "tSupplier", "Pubs", clsArya.DbName, "Where [Date] = '" & txtDate & "'"
            Me.InsertDb "tFacM", "Pubs", clsArya.DbName, "Where [Date] = '" & txtDate & "'"
            Me.InsertDb "tRepfaceditm", "Pubs", clsArya.DbName, "Where [Date] = '" & txtDate & "'"
            Me.InsertDb "tFacD", "Pubs", clsArya.DbName, "Where intserialno in ( select intserialno from Pubs.dbo.tfacm)" ' "Where [Date] = '" & txtDate & "'"
            Me.InsertDb "tFacD2", "Pubs", clsArya.DbName, "where Code in ( select Code from Pubs.dbo.trepfaceditm)" '"Where [Date] = '" & txtDate & "'"
            
            tem_str = dt.Miladi("13" + Mid(txtFile.Text, 1, 2) + "/" + Mid(txtFile.Text, 3, 2) + "/" + Mid(txtFile.Text, 5, 2))
            With cmd
                .ActiveConnection = Conn
                .CommandType = adCmdText
           
                .CommandText = "USE master" & _
                               " EXEC sp_addumpdevice 'disk', 'tmpTotal', '" & Dir1.Path & "\" & Mid(txtFile.Text, 1, 6) & "_" & CStr(cnt) & ".dbk" & "' " & _
                               " BACKUP DATABASE Pubs TO tmpTotal " & _
                               " exec sp_dropdevice 'tmpTotal' "
            End With
            cmd.Execute
            cmd.Cancel
            Me.DropDb "tFacD", "Pubs"
            Me.DropDb "tFacD2", "Pubs"
            Me.DropDb "tFacM", "Pubs"
            Me.DropDb "tRepfaceditm", "Pubs"
            Me.DropDb "tSupplier", "Pubs"
            Me.DropDb "tCust", "Pubs"
    End Select
   ' load frmMsg
    frmMsg.fwlblMsg.Caption = " ›«Ì· Å‘ Ì»«‰ ‰”ŒÂ" & cnt & "  »« „Ê›ﬁÌ  «ÌÃ«œ ê—œÌœ "
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = " «∆Ìœ"
    frmMsg.Show vbModal
End If
Unload frmBackup
Exit Sub
ErrorHandler:
    Unload frmBackup
End Sub

Public Sub ExitForm()
    Unload Me

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


End Sub

Private Sub txtDate_Change()
Select Case Trim(txtDate)
    Case "":
        txtFile.Text = Trim(Mid(clsDate.shamsi(Date), 3, 2)) & Trim(Mid(clsDate.shamsi(Date), 6, 2)) & Trim(Mid(clsDate.shamsi(Date), 9, 2)) & ".fbk"
    Case Else:
        txtFile.Text = Trim(Mid(txtDate, 1, 2)) & Trim(Mid(txtDate, 4, 2)) & Trim(Mid(txtDate, 7, 2)) & ".dbk"
End Select
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    modgl.SetDate KeyDown, txtDate, KeyCode
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    modgl.SetDate KeyPress, txtDate, KeyAscii
End Sub


Public Sub InsertDb(mvarTable As String, dstDatabase As String, srcDatabase As String, Optional Condition As String)
'On Error GoTo ErrorHandler
With cmd
    .ActiveConnection = Conn
    .CommandType = adCmdText
    .CommandText = "Use Master  Select *  Into " & dstDatabase & ".dbo." & mvarTable & " From " & srcDatabase & ".dbo." & mvarTable & " " & Condition
End With
cmd.Execute
cmd.Cancel
Exit Sub
ErrorHandler:
    Resume Next
End Sub

Public Sub DropDb(mvarTable As String, Database As String)
'On Error GoTo ErrorHandler
With cmd
    .ActiveConnection = Conn
    .CommandType = adCmdText
    .CommandText = "Use Master Drop Table " & Database & ".dbo." & mvarTable
End With
cmd.Execute
cmd.Cancel
Exit Sub
ErrorHandler:
    Resume Next
End Sub

