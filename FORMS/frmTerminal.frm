VERSION 5.00
Begin VB.Form frmTerminal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TelnetTTY"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   6
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTerminal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCallerId 
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      Top             =   5400
      Width           =   3375
   End
   Begin VB.TextBox txtInput 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Width           =   7335
   End
   Begin VB.TextBox txtLog 
      Height          =   4695
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   7500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "ÿ—Õ Ê «Ã—« : ‘—ò  «› ÃÌ ¬—Ì«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   7080
      Width           =   7215
   End
   Begin VB.Label Label2 
      Caption         =   "‘„«—Â Â«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   ".........1.........2.........3.........4.........5.........6.........7.........8"
      Height          =   255
      Left            =   8
      TabIndex        =   2
      Top             =   120
      Width           =   7480
   End
End
Attribute VB_Name = "frmTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub txtInput_KeyPress(KeyAscii As Integer)
'    If KeyAscii = Asc(vbCr) Then
'        frmControl.ttcControl.SendData txtInput.Text
'        frmControl.ttcControl.SendData vbCrLf
'        KeyAscii = 0
'        txtLog.Text = txtLog.Text & txtInput.Text & vbCrLf
'        txtLog.SelStart = Len(txtLog.Text)
'        txtInput.Text = ""
'    End If
End Sub

