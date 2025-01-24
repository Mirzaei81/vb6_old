VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"frmRegisterForm.frx":0000
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEscape 
      BackColor       =   &H0000C0C0&
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   4605
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   2040
      TabIndex        =   13
      Top             =   2205
      Width           =   8895
      Begin VB.CommandButton cmdGetData 
         BackColor       =   &H0000C0C0&
         Caption         =   "ê—› ‰ Ê À»  ﬂœ »Â ’Ê—  « Ê„« Ìò"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6600
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblNote 
         Alignment       =   2  'Center
         Caption         =   " « œ—Ì«›  Å«”Œ «“ ‘—ﬂ  «Ì‰ ›—„ —« »«“ ‰êÂœ«—Ìœ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   840
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   " »—«Ì « ’«· »Â „—ﬂ“ «› ÃÌ ¬—Ì« - Œÿ  ·›‰ »Â „Êœ„ Ê’· »«‘œ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   3000
         TabIndex        =   15
         Top             =   1080
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   2040
      TabIndex        =   9
      Top             =   3885
      Width           =   8895
      Begin VB.TextBox txtRegister 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   11
         Top             =   840
         Width           =   4335
      End
      Begin VB.CommandButton CommandButton2 
         BackColor       =   &H0000C0C0&
         Caption         =   "À»  òœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂœ œ—Ì«›  ‘œÂ „Ê«—œ 1 Ì« 2 —« œ— «Ì‰ ﬁ”„  Ê«—œ ﬂ‰Ìœ Ê ﬂ·Ìœ À»  ﬂœ —« ›‘«— œÂÌœ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.Frame Frame_Delegates 
      BackColor       =   &H8000000A&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   10815
      Begin VB.Label Lbl_Limited_Register2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‘‘‘‘‘‘‘‘‘‘‘‘‘‘‘‘‘‘"
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
         Height          =   615
         Left            =   0
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   10575
      End
      Begin VB.Label Lbl_Limited_Register 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‘‘‘‘‘‘‘‘‘‘‘‘‘‘‘‘‘‘"
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
         Height          =   615
         Left            =   0
         TabIndex        =   37
         Top             =   1080
         Visible         =   0   'False
         Width           =   10575
      End
      Begin VB.Label Lbl_Limited2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "œﬁ  ‘Êœ »—«Ì Ãœ« ﬂ—œ‰ »Œ‘ Â«Ì „Œ ·› ﬂœ «—”«·Ì «“ ⁄·«„  + «” ›«œÂ ‘Êœ"
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
         Height          =   615
         Left            =   0
         TabIndex        =   35
         Top             =   1440
         Visible         =   0   'False
         Width           =   10575
      End
      Begin VB.Label Lbl_Limited 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "—« »Â ‘—ﬂ  «› ÃÌ ¬—Ì« 09192671170 - 09192671172 ÅÌ«„ﬂ ﬂ‰Ìœ Ê „‰ Ÿ— œ—Ì«›  Å«”Œ »„«‰Ìœ Ê Ì« »« 88554455  „«” Õ«’· ‰„«ÌÌœ     "
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
         Height          =   735
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   9735
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂœ  Ê·Ìœ ‘œÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Lbl_Takin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "+9821 22263035  —« »«  ·›‰ »Â ‘—ﬂ  ‰ﬂÌ‰ «·ﬂ —Ê‰Ìﬂ Å«”«—ê«œ - Ì« ‰„«Ì‰œÂ ›—Ê‘ - «ÿ·«⁄ œÂÌœ  "
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
         Height          =   855
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   9855
      End
      Begin VB.Label lblGenerateCodeTag2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1560
         TabIndex        =   5
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3480
         TabIndex        =   4
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblLockNo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1800
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblHard2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3720
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lbl_Safir 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "+98511 7243900 - 05117232396 —« »«  ·›‰ »Â ‘—ﬂ  ”›Ì—¬—Ì« - Ì« ‰„«Ì‰œÂ ›—Ê‘ - «ÿ·«⁄ œÂÌœ 09157232396 - 09157232397 "
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
         Height          =   735
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   9975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   2055
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   10815
      Begin VB.Label lblHard 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7320
         TabIndex        =   33
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblCustId 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5280
         TabIndex        =   32
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblspec 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3720
         TabIndex        =   31
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblLockNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1800
         TabIndex        =   30
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7080
         TabIndex        =   29
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5040
         TabIndex        =   28
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3480
         TabIndex        =   27
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1560
         TabIndex        =   26
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblGenerateCodeTag 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "- 3"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   10200
         TabIndex        =   24
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "- 2"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   10200
         TabIndex        =   23
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "- 1"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   10200
         TabIndex        =   22
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ì«  Œÿ  ·›‰ —« »Â „Êœ„ „ ’· ‰„ÊœÂ Ê ò·Ìœ À»  òœ « Ê„« Ìò —« ›‘«— œÂÌœ"
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
         Height          =   495
         Left            =   3000
         TabIndex        =   21
         Top             =   1440
         Width           =   7215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‰„«∆Ìœ Ê „‰ Ÿ— Å«”Œ «“ ÿ—› ¬—Ì« »„«‰Ìœ SMS     Ì« »Â ‘„«—Â 09192671170 - 09192671172 "
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
         Height          =   495
         Left            =   960
         TabIndex        =   20
         Top             =   1080
         Width           =   9135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(88554455 - 88554466 - 88554477 ) Ì« »«  ·›‰ »Â ‘—ﬂ  ¬—Ì«  - ‰„«Ì‰œÂ ›—Ê‘ - «ÿ·«⁄ œÂÌœ  "
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
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   9855
      End
      Begin VB.Label lblGenerateCode 
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂœ  Ê·Ìœ ‘œÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TempCode, TempCode2  As String
Private clsDate As New clsDate
Private TempSpecCust, LockId As String
Private aa As Boolean
Private StrTemp5, StrTemp6 As String
Private f As New FileSystemObject
Private m_cWebService As New cls_WebService

Private Sub cmdGetData_Click()

    Dim WebServiceURL As String
    WebServiceURL = "http://192.168.1.9:1948/samarSecurity"
    Dim ReturnCode As String
    RetValue = m_cWebService.Connect(WebServiceURL)
    If RetValue = 680 Then
        HangUp
        MsgBox ("Œÿ  ·›‰ Ê’· ‰Ì”  -680 ")
     '   AddLog "Œÿ  ·›‰ Ê’· ‰Ì”    -"
        Exit Sub
    ElseIf Status = 676 Then
        HangUp
        MsgBox ("Œÿ „‘€Ê· „Ì »«‘œ  -676")
       ' AddLog Status & "Œÿ „‘€Ê· „Ì »«‘œ  - "
        Exit Sub
    ElseIf Status = 0 Then
        MsgBox ("«— »«ÿ »—ﬁ—«— ‘œ")
       ' AddLog "«— »«ÿ »—ﬁ—«— ‘œ  - "
        cmdGetData.Enabled = False
    Else
        HangUp
        MsgBox (RetValue & "œ—«— »«ÿ »« „Êœ„ „‘ò· ÊÃÊœœ«—œ")
       ' AddLog Status & "œ—«— »«ÿ »« „Êœ„ „‘ò· ÊÃÊœœ«—œ  - "
        Exit Sub
    End If
        
    Dim strCommand As String
    strCommand = lblGenerateCodeTag.Caption & seperator & _
                lblLockNo.Caption & seperator & lblSpec.Caption & seperator & lblCustId.Caption & seperator & Hhhh
    
    modsock.GetCodeRegisterSock strCommand
    ReturnCode = modgl.strSockRecive
    If ReturnCode <> "-109" Then
        If ReturnCode = "-105" Then
            MsgBox "Error:" + ReturnCode + "‘„« «Ã«“Â À»  »—‰«„Â —« ‰œ«—Ìœ"
            Sleep 1000
            cmdGetData.Enabled = True
            Exit Sub
        ElseIf ReturnCode = "-106" Then
            MsgBox "Error:" + ReturnCode + "‘„« «Ã«“Â À»  »—‰«„Â —« ‰œ«—Ìœ"
            Sleep 1000
            cmdGetData.Enabled = True
            Exit Sub
        End If
        txtRegister.Text = ReturnCode
        lblStatus.Caption = "« ’«· »Â „—ò“„Ê›ﬁÌ  ¬„Ì“ »Êœ "
        Sleep 1000
        CommandButton2_Click
    Else
        lblStatus.Caption = "Œÿ« œ—« ’«· »Â „—ﬂ“ "
        lblStatus.ForeColor = vbRed
        Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Sound\notify.wav", True, False)
        Sleep 1000
    End If
    cmdGetData.Enabled = True
End Sub

Private Sub CommandButton2_Click()
Dim tempstring As TextStream
Dim c, D  As String
Dim strTemp, strTemp1, strTemp2, strTemp3, strTemp4, ExpireDate As String
Dim IsFileExist As Boolean
Dim i, LenLockid, IdCustomer As Integer
Dim HardSerialNo As String
Dim RegisterCode  As String
On Error Resume Next

LenLockid = Val(Mid(Me.txtRegister.Text, 8, 1))
If LockId <> Val(Mid(Me.txtRegister.Text, 8, LenLockid + 1)) Then
    MsgBox " Registeration Failed"
    cmdEscape_Click
End If

If clsArya.HardLock = True Or clsArya.LimitedVersion = True Then
    TempCode = TempCode2
Else
End If
RegisterCode = Mid(Me.txtRegister.Text, 3, 3) & Mid(Me.txtRegister.Text, 8)
RegisterCode = left(RegisterCode, Len(RegisterCode) - 2)
If TempCode = RegisterCode Then
    If clsArya.LimitedVersion = True Then
        AppendExpDate
        MsgBox " Successfully Registered"
        Unload Me
'    ElseIf clsArya.LimitedVersion = True And HardLockFlagTrial = True Then
'        Call mdifrm.FWRegistry1.DeleteKeyAll(flwRegLocalMachine, StrTemp5)
'        AppendExpDate
'        MsgBox " Successfully Registered"
'        Unload Me
    Else
         
        Select Case SecurityVersion
            Case 0
                 Select Case Station_IsServer
                    Case True:
                         If left(Me.txtRegister.Text, 2) = "00" Then
    
                            ExpireDate = "Unlimited"
                         Else
                           
                            ExpireDate = "20" & left(Me.txtRegister.Text, 2) & "/" & Mid(Me.txtRegister.Text, 6, 2) & "/" & Right(Me.txtRegister.Text, 2)
                            Dim expDate As Date
                            expDate = ExpireDate
                            If clsArya.MiladiDate = 0 Then
                                 ExpireDate = clsDate.shamsi(expDate)
                            End If
                         End If
                        
                         IsFileExist = f.FileExists(Server_Dir & "\Objectvar2.ini")
                         If IsFileExist = False Then
            
                            f.CreateTextFile Server_Dir & "\Objectvar2.ini"
                         End If
                          
                         For i = clsArya.CustomerId To clsArya.CustomerId + 10
                
                            Set tempstring = f.OpenTextFile(Server_Dir & "\Objectvar2.ini", ForWriting, False, TristateFalse)
                            strTemp = mdifrm.FWEncryption1.Encode(i, 1000)
                            tempstring.WriteLine (strTemp)
                            strTemp1 = mdifrm.FWEncryption1.Encode(ExpireDate, i + 1000)
                            tempstring.WriteLine (strTemp1)
                            strTemp2 = mdifrm.FWEncryption1.Encode("HardLockNo", i + 1000)
                            tempstring.WriteLine (strTemp2)
                            strTemp3 = mdifrm.FWEncryption1.Encode(clsArya.HardLockSerialNo, i + 1000) 'Lock No
                            tempstring.WriteLine (strTemp3)
                            strTemp4 = mdifrm.FWEncryption1.Encode(Hhhh, i + 1000) 'Lock No
                            tempstring.WriteLine (strTemp4)
                
                            tempstring.Close
                
                            Set tempstring = f.OpenTextFile(Server_Dir & "\Objectvar2.ini", ForReading, False, TristateFalse)
                
                            strTemp = tempstring.ReadLine
                            IdCustomer = mdifrm.FWEncryption1.Decode(strTemp, 1000)
                            strTemp = tempstring.ReadLine
                            strTemp1 = mdifrm.FWEncryption1.Decode(strTemp, i + 1000)
                            strTemp = tempstring.ReadLine
                            strTemp2 = mdifrm.FWEncryption1.Decode(strTemp, i + 1000)
                            strTemp = tempstring.ReadLine
                            strTemp3 = mdifrm.FWEncryption1.Decode(strTemp, i + 1000)
                            strTemp = tempstring.ReadLine
                            strTemp4 = mdifrm.FWEncryption1.Decode(strTemp, i + 1000)
                
                            tempstring.Close
                            If IdCustomer = i And ExpireDate = strTemp1 And "HardLockNo" = strTemp2 And clsArya.HardLockSerialNo = strTemp3 And Hhhh = strTemp4 Then
                                Exit For
                            End If
                        Next i
                        Set tempstring = f.OpenTextFile(Server_Dir & "\Objectvar2.ini", ForAppending, False, TristateFalse)
                        For i = 1 To 50
                           strTemp1 = mdifrm.FWEncryption1.Encode(Int((Rnd(1000)) * 1000000 + Rnd(1000) * 1000000000), clsArya.CustomerId + 1000)
                           tempstring.WriteLine (strTemp1)
                
                        Next
                
                        tempstring.Close
                 End Select
            
    ''''            Dim objDisk As FLWDiskFile.IFWDisk
    ''''            i = 0
    ''''            For Each objDisk In mdifrm.FWDisks1.Disks   '
    ''''        ''''      Call cboDisks.AddItem("Drive " & objDisk.Unit & " " & objDisk.TypeName)
    ''''              i = i + 1
    ''''              If InStr(1, objDisk.Unit, "C:\", 1) Then
    ''''                Exit For
    ''''              End If
    ''''            Next
                 Call mdifrm.FWRegistry1.CreateKey(flwRegLocalMachine, StrTemp5)
                 If clsArya.LimitedVersion = False Or (clsArya.LimitedVersion = True And HardLockFlagTrial = True And CustomerRegisterFlag = True) Then
                     StrTemp6 = mdifrm.FWEncryption1.Encode(Hhhh, 2000)
                     If mdifrm.FWRegistry1.SetKeyStr(flwRegLocalMachine, StrTemp5, "String Value", StrTemp6) <> FLWSystem.flwSuccess Then
                         Call MsgBox("Œÿ« œ— À»  «ÿ·«⁄«  - ﬂœ Œÿ« 15  " * vbLf & "Registeration Faile", vbCritical)
                       '  Unload Me
                     End If
                     StrTemp6 = mdifrm.FWEncryption1.Encode(Hhhh, 3000)
                     If mdifrm.FWRegistry1.SetKeyStr(flwRegLocalMachine, StrTemp5, "String Value2", StrTemp6) <> FLWSystem.flwSuccess Then
                         Call MsgBox("Œÿ« œ— À»  «ÿ·«⁄«  - ﬂœ Œÿ« 15  " * vbLf & "Registeration Faile", vbCritical)
                       '  Unload Me
                     End If
                End If
                MsgBox " Successfully Registered"
                Unload Me
            Case 1
         
                 Select Case Station_IsServer
                    Case True:
    '                     If Left(Me.txtRegister.Text, 2) = "00" Then
    '
    '                        ExpireDate = "Unlimited"
    '                     Else
                           
                            ExpireDate = "20" & left(Me.txtRegister.Text, 2) & "/" & Mid(Me.txtRegister.Text, 6, 2) & "/" & Right(Me.txtRegister.Text, 2)
                            expDate = ExpireDate
                            If clsArya.MiladiDate = 0 Then
                                 ExpireDate = clsDate.shamsi(expDate)
                            End If
    '                     End If
                        
                 End Select
            '
            
                Dim strIsServer As String
                strIsServer = IIf(Station_IsServer = True, "1", "0")
                
                Dim strStationsCount As String
                strStationsCount = clsArya.MaxStationNo + clsArya.MaxPocketPcNo
                
                Dim strCommand As String
                strCommand = clsArya.CustomerId & seperator & _
                            clsArya.HardLockSerialNo & seperator & clsArya.StationNo & seperator & Hhhh & seperator & strIsServer _
                            & seperator & ExpireDate & seperator & "HardLockNo" & seperator & strStationsCount
                            
                modsock.SetDefaultServerDataRegisterSock strCommand
                            
                strSockRecive = ""
                If Winsock1.State = sckConnected Then
                    Winsock1.SendData Operations.LogOutStation & seperator & EOS
                End If
                While strSockRecive = ""
                    DoEvents
                Wend
                lblStatus.Caption = " Successfully Registered"
                'MsgBox " Successfully Registered"
                Sleep 1000
                Unload Me
        End Select
    End If
Else
    MsgBox " Registeration Failed"
    cmdEscape_Click
End If

End Sub

Private Sub cmdEscape_Click()
    'modsock.DiscounectSock
    If SecurityVersion = 1 Then
        strSockRecive = ""
        If mdifrm.Winsock1.State = sckConnected Then
            mdifrm.Winsock1.SendData Operations.LogOutStation & seperator & EOS
        End If
        Dim jjj As Long
        While strSockRecive = ""
            DoEvents
            jjj = jjj + 1
            If jjj = 200000 Then
                End
            End If
        Wend
    End If
    End
End Sub

Private Sub Form_Activate()
    
    If clsArya.HardLock = True Or clsArya.LimitedVersion = True Then
        cmdGetData.Visible = False
        lblStatus.Visible = False
        LblNote.Visible = True
        Frame_Delegates.Visible = True
        Frame1.Visible = False
        If strDelegate = "11" Then
            lbl_Safir.Visible = True
        ElseIf strDelegate = "24" Then
            Lbl_Takin.Visible = True
        Else
            Lbl_Limited.Visible = True
            Lbl_Limited2.Visible = True
        End If
    Else
        Frame_Delegates.Visible = False
        Frame1.Visible = True
    End If
    
    StrTemp5 = mdifrm.FWEncryption1.Decode("Õ∞`Âr24∆°◊vÒÄ—W„ÿV$3¥ã˝ıÜîJı\˘`", 2000)  '  "Software\Microsoft\Visual Program"
    
    Randomize
    Me.lblGenerateCodeTag.Caption = Int((Rnd(1)) * 100000) ' & "-" & ClsStation.CustomerId
    Me.lblGenerateCodeTag2.Caption = Int((Rnd(1)) * 100000) ' & "-" & ClsStation.CustomerId
    
    aa = clsArya.DemoVersion
    CustSpecFind aa
    aa = clsArya.LimitedVersion
    CustSpecFind aa
    aa = clsArya.TrialVer
    CustSpecFind aa
    aa = clsArya.SoftLock
    CustSpecFind aa
    aa = clsArya.HardLock
    CustSpecFind aa
    
    
       Me.lblLockNo.Caption = clsArya.HardLockSerialNo
       Me.lblLockNo2.Caption = clsArya.HardLockSerialNo
''''    If Len(clsArya.HardLockSerialNo) = 15 Then
''''       Me.lblLockNo.Caption = Mid(clsArya.HardLockSerialNo, 11, 4)
''''    Else
''''       Me.lblLockNo.Caption = Right(clsArya.HardLockSerialNo, 4)
''''    End If
    Me.lblSpec.Caption = TempSpecCust & 4         ' New Security
    
    Me.lblSpec.Caption = Me.lblSpec.Caption & clsArya.StationNo            '
    
    If Station_IsServer = True Then
        Me.lblSpec.Caption = Me.lblSpec.Caption & 1           '
    Else
        Me.lblSpec.Caption = Me.lblSpec.Caption & 0           '
    End If
    
    If f.FileExists(Server_Dir & "\Objectvar2.ini") = True Then
        Me.lblSpec.Caption = Me.lblSpec.Caption & 1           '
    Else
        Me.lblSpec.Caption = Me.lblSpec.Caption & 0           '
    End If
    
    If mdifrm.FWRegistry1.KeyExists(flwRegLocalMachine, StrTemp5) = True Then
        Me.lblSpec.Caption = Me.lblSpec.Caption & 1           '
    Else
        Me.lblSpec.Caption = Me.lblSpec.Caption & 0           '
    End If
    
''''    Me.lblCustId.Caption = Mid(StringExeMaker, 1, 8)
    Me.lblCustId.Caption = strDelegate & strCategory & Format(clsArya.CustomerId, "00000") & intVersion
    Dim aqq As String
    Dim secureBytes() As Byte
    Dim index As Integer
    
    secureBytes = Hhhh
    
''    aqq = ""
''    For Index = LBound(secureBytes) To UBound(secureBytes)
''        secureBytes(Index) = secureBytes(Index) Xor 7
''        If secureBytes(Index) > 10 Then
''            aqq = aqq & Chr$(secureBytes(Index))
''        End If
''    Next Index
    
'    Me.lblHard.Caption = Val(aqq)
    
    Me.lblHard.Caption = Hhhh
    Me.lblHard2.Caption = Me.lblSpec.Caption

    TempCode = Int(lblGenerateCodeTag.Caption * 33.789) + 22456
    
    ' For Limited Version
    If clsArya.LimitedVersion = True And HardLockFlagTrial = True And CustomerRegisterFlag = True Then
        Me.lblHard2.Caption = "123"
        TempCode2 = Int(lblGenerateCodeTag2.Caption * 39.437) + 34521
        Lbl_Limited.Visible = False
        Lbl_Limited2.Visible = False
'        Lbl_Limited_Register.Visible = True
'        Lbl_Limited_Register2.Visible = True
'        Lbl_Limited_Register = "»—«Ì «” ›«œÂ «“ «Ì‰ ”Ì” „ „‘Œ’«  œ«—‰œÂ ¬‰ »«Ìœ œ— ”Ì” „ „‘ —Ì«‰ ‘—ﬂ  ¬˙—Ì« ”„— À»  ê—œœ"
'        Lbl_Limited_Register2 = "»—«Ì À»  „‘Œ’«  »«  ·›‰ Â«Ì ‘—ﬂ  Ê Ì« ‰„«Ì‰ê«‰ ›—Ê‘ ¬‰  „«” Õ«’· ›—„«∆Ìœ  "
    Else   '
        TempCode2 = Int(lblGenerateCodeTag2.Caption * 42.437) + 27462 + Val(Me.lblHard2.Caption)
    End If

'    Me.lblHard.Caption = Hhhh
    
'    secureBytes = aqq
'    aqq = ""
'    For index = LBound(secureBytes) To UBound(secureBytes)
'        secureBytes(index) = secureBytes(index) Xor 7
'        If secureBytes(index) > 10 Then
'            aqq = aqq & Chr$(secureBytes(index))
'        End If
'    Next index
'
    If clsArya.HardLock = True Or clsArya.LimitedVersion = True Then
        If clsArya.LimitedVersion = True And HardLockFlagTrial = True And CustomerRegisterFlag = True Then
            LockId = (Val(Right(Me.lblLockNo.Caption, 4)) + 97) * Val(left(lblGenerateCodeTag2.Caption, 1))
        Else
            LockId = (Val(Right(Me.lblLockNo.Caption, 4))) * Val(left(lblGenerateCodeTag2.Caption, 1))
        End If
    Else
        LockId = (Val(Right(Me.lblLockNo.Caption, 4)) + Val(left(Format(clsArya.CustomerId, "00000"), 3))) * 3
    End If
    LockId = Len(LockId) & LockId
    
    TempCode = left(TempCode, 3) & CStr(LockId) & Mid(TempCode, 4)
    TempCode2 = left(TempCode2, 3) & CStr(LockId) & Mid(TempCode2, 4)

    lblStatus.Caption = " »—«Ì « ’«· »Â „—ﬂ“ «› ÃÌ ¬—Ì« -  Œÿ  ·›‰ »Â „Êœ„ Ê’· »«‘œ"
    
'    If SecurityVersion = 0 Then
'        cmdGetData.Enabled = False
'        Label4.Visible = False
'        Label7.Visible = False
'    End If

End Sub

Private Function CustSpecFind(index As Boolean)
   Select Case index
   
      Case True
          TempSpecCust = TempSpecCust & "1"
   
      Case False
          TempSpecCust = TempSpecCust & "0"
   
   End Select
End Function






