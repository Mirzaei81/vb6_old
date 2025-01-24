VERSION 5.00
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmMsg 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   -45
      Width           =   6195
      Begin FLWCtrls.FWLabel3D fwlblMsg 
         Height          =   1935
         Left            =   120
         Top             =   240
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   3413
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Deep            =   100
         ForeColor1      =   16761024
         ForeColor2      =   0
         BackColor       =   8421631
         Caption         =   "¬Ì« ‘„« „Ì ŒÊ«ÂÌœ ﬂÂ œ— »—‰«„Â ÅÌ€«„ »«‘œ ø"
         Alignment       =   1
      End
      Begin FLWCtrls.FWButton fwBtn 
         Height          =   525
         Index           =   0
         Left            =   1800
         TabIndex        =   1
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   926
         Caption         =   "ﬁ»Ê·"
         BackColor       =   -2147483633
         FontName        =   "B Traffic"
         FontBold        =   -1  'True
         FontSize        =   12
         Alignment       =   1
      End
      Begin FLWCtrls.FWButton fwBtn 
         Height          =   525
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   926
         ButtonType      =   1
         Caption         =   "Œ—ÊÃ"
         BackColor       =   -2147483633
         FontName        =   "B Traffic"
         FontBold        =   -1  'True
         FontSize        =   12
         Alignment       =   1
      End
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    CenterCenterinSecondScreen Me
End Sub

Private Sub fwBtn_Click(Index As Integer)
Dim i As Double
    frmMsg.BackColor = &HC0C0C0
''''    Select Case OptionSelect(0)
''''        Case True:
''''            modgl.mvarMsgSelect = 0
''''    End Select
''''    Select Case OptionSelect(1)
''''        Case True:
''''            modgl.mvarMsgSelect = 1
''''    End Select
    Select Case Index
        Case 0:
            modgl.mvarMsgIdx = vbYes
        Case 1:
            modgl.mvarMsgIdx = vbNo
    End Select
''''    SendKeys "{Left}", True
    Unload Me
    
End Sub

Private Sub fwBtn_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
   modgl.mvarMsgIdx = vbNo
Else
     modgl.mvarMsgIdx = vbYes
End If
Unload Me

End Sub
