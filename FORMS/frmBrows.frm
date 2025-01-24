VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmBrowser 
   Caption         =   "›«Ì· —«Â‰„«"
   ClientHeight    =   6630
   ClientLeft      =   3060
   ClientTop       =   1650
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "Nazanin"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   5310
      Left            =   75
      TabIndex        =   8
      Top             =   1185
      Width           =   11685
      ExtentX         =   20611
      ExtentY         =   9366
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   11865
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11865
      Begin VB.CommandButton cmdFgarya 
         Caption         =   "”«Ì  ¬—Ì«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton CmdGoogle 
         Caption         =   "Google"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "ﬁ»·Ì"
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdForward 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         ToolTipText     =   "»⁄œÌ"
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         ToolTipText     =   " Êﬁ›"
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cboAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   360
         Width           =   4515
      End
      Begin VB.Label lblAddress 
         Caption         =   "&Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Tag             =   "&Address:"
         Top             =   0
         Width           =   3075
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   9480
         Picture         =   "frmBrows.frx":A4C2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2370
      End
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6180
      Top             =   1500
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

Private Sub cmdBack_Click()
    On Error Resume Next
    timTimer.Enabled = True
 
    brwWebBrowser1.GoBack
End Sub

Public Sub cmdFgarya_Click()
    Dim st As String
    If strDelegate = "56" Then st = "http://www.MoeinReklam.com" Else st = "http://www.fgarya.com"
    cboAddress.Text = st
    cboAddress_Click
End Sub
Public Sub CmdGoogle_Click()
    cboAddress.Text = "http://www.Google.com"
    cboAddress_Click

End Sub

Private Sub CmdForward_Click()
''''    On Error Resume Next
''''    timTimer.Enabled = True
''''
''''    brwWebBrowser1.GoForward
End Sub
Public Sub ExitForm()

    Unload Me
    
End Sub


Public Sub Printing()
    Me.brwWebBrowser.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub


Private Sub cmdStop_Click()
''''    On Error Resume Next
''''
''''    timTimer.Enabled = False
''''    brwWebBrowser1.Stop
''''    Me.Caption = brwWebBrowser1.LocationName

End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Load()
    On Error Resume Next
  '  Me.Show
    tbToolBar.Refresh
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

    Form_Resize

    cboAddress.Move 50, lblAddress.top + lblAddress.Height + 15


    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser1.Navigate StartingAddress
    End If
    Dim filetemp As New FileSystemObject
    Dim formName As String
    If Not IsHelp Then
       cboAddress.AddItem App.Path & "/help/index.html"
       cboAddress.Text = App.Path & "/help/index.html"
    Else
        formName = Mid(VarActForm, 4) & ".htm"
        If filetemp.FileExists(App.Path & "/help/" & formName) = True Then
            cboAddress.Text = App.Path & "/help/" & formName
        Else
            cboAddress.Text = App.Path & "/help/index.html"
        End If
    End If
    
    Dim st As String
    If strDelegate = "56" Then st = "www.MoeinReklam.com" Else st = "www.fgarya.com"
    cmdFgarya.Caption = st
    
    If strDelegate <> "56" Then
        cboAddress.AddItem "www.Fgarya.com"
    '    m_IpAddresses.Add "164.138.16.20"
    
        cboAddress.AddItem "www.AryaSmsPanel.ir"
        cboAddress.AddItem "www.SafirArya.ir"
    '    m_IpAddresses.Add "87.107.121.52"
    End If
    cboAddress.AddItem "www.google.com"
'    m_IpAddresses.Add "173.194.70.100"

    cboAddress.AddItem "www.Yahoo.com"
'    m_IpAddresses.Add "98.138.253.109"

    cboAddress.AddItem "msdn.microsoft.com"
'    m_IpAddresses.Add "207.46.248.109"
    cboAddress_Click


End Sub

Private Sub brwWebBrowser1_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser1.LocationName
End Sub


Private Sub brwWebBrowser1_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser1.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser1.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser1.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub


Public Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
 '   timTimer.Enabled = True
    brwWebBrowser1.Navigate cboAddress.Text
End Sub


Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub
''''Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
''''    If formloadFlag = True Then
''''        SaveSetting strMainKey, Me.Name, "Height", Me.Height
''''        SaveSetting strMainKey, Me.Name, "Width", Me.Width
''''    End If
''''
''''End Sub


Private Sub Form_Resize()
    On Error Resume Next
'    cboAddress.Width = Me.ScaleWidth - 100
    brwWebBrowser1.Width = Me.ScaleWidth
    brwWebBrowser1.Height = Me.ScaleHeight - 1000 '- (cmdBack.Top + cmdBack.Height) - 100
    brwWebBrowser1.top = cmdBack.top + cmdBack.Height + 50
    brwWebBrowser1.left = 0
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If
    Image1.left = Me.Width - Image1.Width - 250
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top
    VarActForm = ""
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser1.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser1.LocationName
    Else
        Me.Caption = "Working... "
    End If
End Sub


