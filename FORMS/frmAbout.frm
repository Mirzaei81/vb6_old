VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   ClientHeight    =   8430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   Enabled         =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame_Moein 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MoeinReklam@gmail.com"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   1200
         TabIndex        =   27
         Top             =   7080
         Width           =   3975
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "07708615501 - 07708615502 - 07480151660"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   0
         TabIndex        =   26
         Top             =   6000
         Width           =   6615
      End
      Begin VB.Image Image_Moein 
         Height          =   1635
         Left            =   720
         Picture         =   "frmAbout.frx":B16C
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   2100
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Restaurant System"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1320
         TabIndex        =   25
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Image Image2 
         Height          =   1800
         Index           =   1
         Left            =   3960
         Picture         =   "frmAbout.frx":9757E
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1800
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "In Order Of Moein Reklam Co"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   735
         Left            =   480
         TabIndex        =   24
         Top             =   5400
         Width           =   5535
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "www.MoeinReklam.com"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   1200
         TabIndex        =   23
         Top             =   6600
         Width           =   3975
      End
   End
   Begin VB.Image Image_Dealer 
      Height          =   1215
      Left            =   120
      Top             =   5760
      Width           =   6615
   End
   Begin VB.Label lblDealer 
      Alignment       =   2  'Center
      Caption         =   "‰„«Ì‰œÂ ›—Ê‘ :                            "
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   6120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Ê·Ìœ ﬂ‰‰œÂ ”Ì” „ Â«Ì  Œ’’Ì ›—Ê‘ "
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   480
      TabIndex        =   20
      Top             =   2520
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Image Image2 
      Height          =   1815
      Index           =   0
      Left            =   1560
      Picture         =   "frmAbout.frx":99646
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.Image Image5 
      Height          =   840
      Left            =   960
      Picture         =   "frmAbout.frx":9E28C
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1425
   End
   Begin VB.Label w 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "»—«Ì »Â—Â „‰œÌ «“ ¬Œ—Ì‰ ﬁ«»·Ì  Â« Ê ‰”ŒÂ Â«Ì ‰—„ «›“«—Ì »« ‰„«Ì‰œê«‰ ›—Ê‘  „«” Õ«’· ‰„«∆Ìœ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   735
      Left            =   240
      TabIndex        =   19
      Top             =   7200
      Width           =   5655
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”Ì” „ —” Ê—«‰Ì "
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3120
      TabIndex        =   18
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1560
      Index           =   0
      Left            =   2280
      Picture         =   "frmAbout.frx":9E834
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1680
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Arya"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   300
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   600
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¬—Ì«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6120
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÃÌ"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5880
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«›"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰: 88554455-88554466-88554477-88554488"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   615
      Left            =   -360
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¬—Ì«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÃÌ"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«›"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Arya"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«—«∆Â «‰Ê«⁄ ”Œ  «›“«— Ê ·Ê«“„ Ã«‰»Ì ›—Ê‘ "
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WWW.Fgarya.Com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Info@Fgarya.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "œ«—«Ì êÊ«ÂÌ — »Â »‰œÌ «“ ‘Ê—«Ì⁄«·Ì «‰›Ê—„« Ìò ò‘Ê—"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Ê·Ìœ Ê«—«∆Â ‰—„ «›“«—Â«Ì ”›«—‘Ì"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   5535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CenterCenter Me
    Me.Left = Me.Left - (frmGroupMenu.Width / 2): Me.Top = 300
    Me.BackColor = mdifrm.BackColor
    Dim filetemp As New FileSystemObject
    Dim LogoFile As String
    LogoFile = App.Path & "\Image\Logo_Dealer.gif"
    If filetemp.FileExists(LogoFile) Then
        Image_Dealer.Picture = LoadPicture(LogoFile)
    Else
        LogoFile = App.Path & "\Image\Logo_Dealer.jpg"
        If filetemp.FileExists(LogoFile) Then
            Image_Dealer.Picture = LoadPicture(LogoFile)
        Else
            lblDealer.Visible = True
        End If
    End If
    Frame_Moein.BackColor = mdifrm.BackColor
    If strDelegate = "56" Then Frame_Moein.Visible = True Else Frame_Moein.Visible = False
End Sub

