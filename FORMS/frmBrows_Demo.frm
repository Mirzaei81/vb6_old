VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmBrowser_Demo 
   Caption         =   "ÝÇíá ÑÇåäãÇ"
   ClientHeight    =   6135
   ClientLeft      =   3060
   ClientTop       =   1650
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Nazanin"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrows_Demo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   6720
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
      TabIndex        =   6
      ToolTipText     =   "ÊæÞÝ"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
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
      TabIndex        =   5
      ToolTipText     =   "ÈÚÏí"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
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
      TabIndex        =   4
      ToolTipText     =   "ÞÈáí"
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6180
      Top             =   1500
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
      ScaleWidth      =   6720
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   6720
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
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   120
         Width           =   3795
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   4320
         Picture         =   "frmBrows_Demo.frx":A4C2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2370
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
         Left            =   0
         TabIndex        =   1
         Tag             =   "&Address:"
         Top             =   0
         Width           =   3075
      End
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
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
Attribute VB_Name = "frmBrowser_Demo"
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

    cboAddress.Move 50, lblAddress.top + lblAddress.Height + 15

    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If
    If Not IsHelp Then
       cboAddress.Text = App.Path & "/rep.htm"
    Else
'        If mvarCategory = Restaurant Or mvarCategory = Club Then
'             cboAddress.Text = App.Path & "/help/help_V26.htm"
'        ElseIf mvarCategory = shop Or mvarCategory = Taavoni Or mvarCategory = Ghanadi Then
             cboAddress.Text = App.Path & "/help/demo_V26_Help.htm"
'        ElseIf mvarCategory = Beauty Then
'             cboAddress.Text = App.Path & "/help/help_Normal.htm"
'        Else
'             cboAddress.Text = App.Path & "/help/help_Normal.htm"
'        End If
    End If
    formloadFlag = False
    Me.left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 5000 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
      '  If Me.Height < 1000 Then Me.Height = 2000
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 5000 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    If Me.left < 0 Then Me.left = 0
    If Me.top < 0 Then Me.top = 0
    If Me.top > Me.ScaleHeight Then Me.top = 0

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
    brwWebBrowser.Navigate cboAddress.Text
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
    brwWebBrowser.top = cmdBack.top + cmdBack.Height + 50
    End If
    brwWebBrowser.left = Me.ScaleLeft   '0
    
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Working... "
    End If
End Sub


