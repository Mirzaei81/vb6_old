VERSION 5.00
Begin VB.Form frmDisMsg 
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1410
      Top             =   360
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1665
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmDisMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'    Me.Left = (Screen.Width - Me.Width) / 2
'    Me.Top = (Screen.Height - Me.Height) / 2
    If clsStation.SoundAlarm = True And VarActForm <> "FrmLogin" Then
'        Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Sound\winAquariumError.wav", True, False)
        Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Sound\Notify.wav", True, False)
    End If
    CenterCenterinSecondScreen Me
    
    'frmDisMsg.lblMessage.Font = "traffic"
    frmDisMsg.lblMessage.FontBold = True
    frmDisMsg.lblMessage.FontSize = 12

End Sub



Private Sub Timer1_Timer()

    Unload Me
    
End Sub
