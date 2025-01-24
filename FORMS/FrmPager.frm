VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmPager 
   Caption         =   $"FrmPager.frx":0000
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17115
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   17115
   StartUpPosition =   3  'Windows Default
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   120
      OleObjectBlob   =   "FrmPager.frx":00C2
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "B Mitra"
         Size            =   48
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1590
      Left            =   0
      TabIndex        =   4
      Text            =   "ãËÇá "
      Top             =   8160
      Width           =   17055
   End
   Begin FLWCtrls.FWScrollText FWScrollText1 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   9840
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   1085
      Caption         =   "ØÑÍ æ ÇÌÑÇÁ : ÔÑ˜Ê ãåäÏÓí Ýä ÂæÑÓÊÑ ÂÑíÇ   ÊáÝä      88554488  -  88554477  - 88554466 - 88554455  9821    www.FGArya.Com "
      BorderStyle     =   9
      ForeColor       =   16448
      BackColor       =   49344
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   14.25
      Interval        =   10
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   80.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3135
      Left            =   840
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4920
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   150
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   4695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "124"
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   249.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   8055
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "frmPager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Text1.Width = Me.Width
    FWScrollText1.Width = Me.Width
End Sub

Private Sub Form_Load()
    
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
    
    Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = ""
    Text1.Text = GetSetting(strMainKey, Me.Name, "Text1")
    Label1.Caption = GetSetting(strMainKey, Me.Name, "Label1")
    Label2.Caption = GetSetting(strMainKey, Me.Name, "Label2")
    Label3.Caption = GetSetting(strMainKey, Me.Name, "Label3")

End Sub

Public Sub UpdateNumber()
    On Error Resume Next
'    FWLabel1.Font.Size = 250
'    FWLabel1.Caption = PagerNo

    
    If PagerNo <> Val(Label1.Caption) And PagerNo <> Val(Label2.Caption) And PagerNo <> Val(Label3.Caption) Then
        Label3.Font.Size = 120
        Label3.Caption = Label2.Caption
        Label2.Font.Size = 180
        Label2.Caption = Label1.Caption
        Label1.Font.Size = 270
        Label1.Caption = PagerNo '& 1
        Text1.TabStop = Falset
        If DebugMode = True Then frmPager.SetFocus
        SaveSetting strMainKey, Me.Name, "Label1", Label1.Caption
        SaveSetting strMainKey, Me.Name, "Label2", Label2.Caption
        SaveSetting strMainKey, Me.Name, "Label3", Label3.Caption
    End If
End Sub

Private Sub Form_Resize()
    
    Text1.Width = Me.Width
    FWScrollText1.Width = Me.Width

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

Private Sub Text1_Change()
        
    SaveSetting strMainKey, Me.Name, "Text1", Text1.Text
    
End Sub
