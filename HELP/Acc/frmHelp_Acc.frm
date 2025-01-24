VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmHelp_Acc 
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   6705
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   6705
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   120
         Width           =   3795
      End
      Begin VB.Label lblAddress 
         Caption         =   "&Address:"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Tag             =   "&Address:"
         Top             =   0
         Width           =   3075
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   4320
         Picture         =   "frmHelp_Acc.frx":0000
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
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "ÞÈáí"
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdForward 
      Caption         =   ">"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "ÈÚÏí"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "X"
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      ToolTipText     =   "ÊæÞÝ"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   4575
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   6600
      ExtentX         =   11642
      ExtentY         =   8070
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmHelp_Acc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

Private Sub cmdBack_Click()
    On Error Resume Next
    timTimer.Enabled = True
 
    brwWebBrowser.GoBack
End Sub

Private Sub CmdForward_Click()
''''    On Error Resume Next
''''    timTimer.Enabled = True
''''
''''    brwWebBrowser.GoForward
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
''''    brwWebBrowser.Stop
''''    Me.Caption = brwWebBrowser.LocationName

End Sub

Private Sub Form_Activate()
    Form_Resize
End Sub

Private Sub Form_Load()
    On Error Resume Next
  '  Me.Show
    tbToolBar.Refresh

    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15

    If Len(StartingAddress) > 0 Then
        cboAddress.text = StartingAddress
        cboAddress.AddItem cboAddress.text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If
    cboAddress.text = App.Path & "/help/Acc/Help_Acc.htm"
    End If
    formloadFlag = False
    Me.Left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 0 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
      '  If Me.Height < 1000 Then Me.Height = 2000
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 0 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.Top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    formloadFlag = True
    cboAddress_Click


End Sub

Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub


Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub


Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
 '   timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.text
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
    If Me.ScaleWidth > 100 Then
        'cboAddress.Width = Me.ScaleWidth - 100
        brwWebBrowser.Width = Me.ScaleWidth
    End If
    If Me.ScaleHeight > 100 Then
    brwWebBrowser.Height = Me.ScaleHeight '- (cmdBack.Top + cmdBack.Height) - 100
    brwWebBrowser.Top = cmdBack.Top + cmdBack.Height + 50
    End If
    brwWebBrowser.Left = Me.ScaleLeft   '0
    
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Working... "
    End If
End Sub



