VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmGoodMenuSetting 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   Icon            =   "frmGoodMenuSetting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   9435
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   6
      Tab             =   2
      TabHeight       =   520
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ÈÎÔ 1"
      TabPicture(0)   =   "frmGoodMenuSetting.frx":A4C2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame14(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ÈÎÔ 2"
      TabPicture(1)   =   "frmGoodMenuSetting.frx":A4DE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame14(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ÈÎÔ 3"
      TabPicture(2)   =   "frmGoodMenuSetting.frx":A4FA
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame14(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "ÈÎÔ 4"
      TabPicture(3)   =   "frmGoodMenuSetting.frx":A516
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame14(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "ÈÎÔ 5"
      TabPicture(4)   =   "frmGoodMenuSetting.frx":A532
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame14(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "ÈÎÔ 6"
      TabPicture(5)   =   "frmGoodMenuSetting.frx":A54E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame14(5)"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame14 
         Caption         =   "ÈÎÔ 6"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   7215
         Index           =   5
         Left            =   -74280
         RightToLeft     =   -1  'True
         TabIndex        =   153
         Top             =   840
         Width           =   6255
         Begin VB.TextBox TxtHeadeTitr 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Titr"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   5
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   193
            Text            =   "ÇÕáí 5"
            Top             =   240
            Width           =   1845
         End
         Begin VB.CheckBox ChkDescription 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÊæÖíÍÇÊ ˜ÇáÇ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   3840
            TabIndex        =   168
            Top             =   6480
            Width           =   2055
         End
         Begin VB.CheckBox ChkPicture 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ú˜Ó"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   4440
            TabIndex        =   167
            Top             =   6000
            Width           =   1455
         End
         Begin VB.TextBox TxtFee2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   5
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   166
            Text            =   "Ýí ÈíÑæä"
            Top             =   5640
            Width           =   1845
         End
         Begin VB.CheckBox ChkFee2 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ýí Ïæã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   4440
            TabIndex        =   165
            Top             =   5520
            Width           =   1455
         End
         Begin VB.TextBox TxtFee 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   5
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   164
            Text            =   "Ýí ÓÇáä"
            Top             =   5160
            Width           =   1845
         End
         Begin VB.CheckBox ChkFee1 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ýí Çæá"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   4440
            TabIndex        =   163
            Top             =   5040
            Width           =   1455
         End
         Begin VB.TextBox TxtName 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   5
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Text            =   "äÇã ˜ÇáÇ"
            Top             =   4680
            Width           =   1845
         End
         Begin VB.CheckBox ChkName 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ äÇã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   4440
            TabIndex        =   161
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox TxtRow 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   5
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   160
            Text            =   "ÑÏíÝ"
            Top             =   4200
            Width           =   1845
         End
         Begin VB.CheckBox ChkRow 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÑÏíÝ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   4440
            TabIndex        =   159
            Top             =   4080
            Width           =   1455
         End
         Begin VB.TextBox TxtFontGrid 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   5
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Text            =   "Titr"
            Top             =   2520
            Width           =   1845
         End
         Begin VB.TextBox TxtFontsizeGrid 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   5
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Text            =   "16"
            Top             =   3000
            Width           =   645
         End
         Begin VB.TextBox TxtHeaderFont 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   5
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   156
            Text            =   "Titr"
            Top             =   720
            Width           =   1845
         End
         Begin VB.TextBox TxtHeaderSize 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   5
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Text            =   "16"
            Top             =   1200
            Width           =   645
         End
         Begin VB.CheckBox ChkViewSegment 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÈÎÔ "
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   4560
            TabIndex        =   154
            Top             =   360
            Width           =   1455
         End
         Begin FLWCtrls.FWComboColor FWComboColorHeader 
            Height          =   315
            Index           =   5
            Left            =   600
            TabIndex        =   169
            Tag             =   "&H00000080&"
            Top             =   1680
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
         End
         Begin FLWCtrls.FWComboColor FWComboColorGrid 
            Height          =   315
            Index           =   5
            Left            =   600
            TabIndex        =   170
            Tag             =   "&H00000080&"
            Top             =   3600
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÑ ÊíÊÑ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   11
            Left            =   3000
            TabIndex        =   194
            Top             =   240
            Width           =   975
         End
         Begin VB.Line Line6 
            X1              =   0
            X2              =   6240
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   2760
            TabIndex        =   182
            Top             =   5640
            Width           =   1095
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   2760
            TabIndex        =   181
            Top             =   5160
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   2760
            TabIndex        =   180
            Top             =   4680
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   2760
            TabIndex        =   179
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "äæÚ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   2760
            TabIndex        =   178
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÇíÒ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   2760
            TabIndex        =   177
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Ñä Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   2760
            TabIndex        =   176
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "ÌÏæá ˜ÇáÇåÇ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   5
            Left            =   4080
            TabIndex        =   175
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "äæÚ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   2880
            TabIndex        =   174
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÇíÒ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   2880
            TabIndex        =   173
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ñä Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   2880
            TabIndex        =   172
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   5
            Left            =   4200
            TabIndex        =   171
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "ÈÎÔ 5"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   7215
         Index           =   4
         Left            =   -74280
         RightToLeft     =   -1  'True
         TabIndex        =   123
         Top             =   840
         Width           =   6255
         Begin VB.TextBox TxtHeadeTitr 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Titr"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   191
            Text            =   "ÇÕáí 5"
            Top             =   240
            Width           =   1845
         End
         Begin VB.CheckBox ChkDescription 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÊæÖíÍÇÊ ˜ÇáÇ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   3840
            TabIndex        =   138
            Top             =   6480
            Width           =   2055
         End
         Begin VB.CheckBox ChkPicture 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ú˜Ó"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   4440
            TabIndex        =   137
            Top             =   6000
            Width           =   1455
         End
         Begin VB.TextBox TxtFee2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   136
            Text            =   "Ýí ÈíÑæä"
            Top             =   5640
            Width           =   1845
         End
         Begin VB.CheckBox ChkFee2 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ýí Ïæã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   4440
            TabIndex        =   135
            Top             =   5520
            Width           =   1455
         End
         Begin VB.TextBox TxtFee 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Text            =   "Ýí ÓÇáä"
            Top             =   5160
            Width           =   1845
         End
         Begin VB.CheckBox ChkFee1 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ýí Çæá"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   4440
            TabIndex        =   133
            Top             =   5040
            Width           =   1455
         End
         Begin VB.TextBox TxtName 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Text            =   "äÇã ˜ÇáÇ"
            Top             =   4680
            Width           =   1845
         End
         Begin VB.CheckBox ChkName 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ äÇã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   4440
            TabIndex        =   131
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox TxtRow 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   130
            Text            =   "ÑÏíÝ"
            Top             =   4200
            Width           =   1845
         End
         Begin VB.CheckBox ChkRow 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÑÏíÝ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   4440
            TabIndex        =   129
            Top             =   4080
            Width           =   1455
         End
         Begin VB.TextBox TxtFontGrid 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Text            =   "Titr"
            Top             =   2520
            Width           =   1845
         End
         Begin VB.TextBox TxtFontsizeGrid 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Text            =   "16"
            Top             =   3000
            Width           =   645
         End
         Begin VB.TextBox TxtHeaderFont 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Text            =   "Titr"
            Top             =   720
            Width           =   1845
         End
         Begin VB.TextBox TxtHeaderSize 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Text            =   "16"
            Top             =   1200
            Width           =   645
         End
         Begin VB.CheckBox ChkViewSegment 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÈÎÔ "
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   4560
            TabIndex        =   124
            Top             =   360
            Width           =   1455
         End
         Begin FLWCtrls.FWComboColor FWComboColorHeader 
            Height          =   315
            Index           =   4
            Left            =   600
            TabIndex        =   139
            Tag             =   "&H00000080&"
            Top             =   1680
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
         End
         Begin FLWCtrls.FWComboColor FWComboColorGrid 
            Height          =   315
            Index           =   4
            Left            =   600
            TabIndex        =   140
            Tag             =   "&H00000080&"
            Top             =   3600
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÑ ÊíÊÑ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   10
            Left            =   3000
            TabIndex        =   192
            Top             =   240
            Width           =   975
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   6240
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   2760
            TabIndex        =   152
            Top             =   5640
            Width           =   1095
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   2760
            TabIndex        =   151
            Top             =   5160
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   2760
            TabIndex        =   150
            Top             =   4680
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   2760
            TabIndex        =   149
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "äæÚ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   2760
            TabIndex        =   148
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÇíÒ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   2760
            TabIndex        =   147
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Ñä Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   2760
            TabIndex        =   146
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "ÌÏæá ˜ÇáÇåÇ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   4
            Left            =   4440
            TabIndex        =   145
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "äæÚ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   3120
            TabIndex        =   144
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÇíÒ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   3120
            TabIndex        =   143
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ñä Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   3120
            TabIndex        =   142
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   4
            Left            =   4440
            TabIndex        =   141
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "ÈÎÔ 4"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   7215
         Index           =   3
         Left            =   -74280
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   840
         Width           =   6255
         Begin VB.TextBox TxtHeadeTitr 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Titr"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   3
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   189
            Text            =   "ÇÕáí 4"
            Top             =   240
            Width           =   1845
         End
         Begin VB.CheckBox ChkDescription 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÊæÖíÍÇÊ ˜ÇáÇ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   3840
            TabIndex        =   108
            Top             =   6480
            Width           =   2055
         End
         Begin VB.CheckBox ChkPicture 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ú˜Ó"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   4440
            TabIndex        =   107
            Top             =   6000
            Width           =   1455
         End
         Begin VB.TextBox TxtFee2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   3
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Text            =   "Ýí ÈíÑæä"
            Top             =   5640
            Width           =   1845
         End
         Begin VB.CheckBox ChkFee2 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ýí Ïæã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   4440
            TabIndex        =   105
            Top             =   5520
            Width           =   1455
         End
         Begin VB.TextBox TxtFee 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   3
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Text            =   "Ýí ÓÇáä"
            Top             =   5160
            Width           =   1845
         End
         Begin VB.CheckBox ChkFee1 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ýí Çæá"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   4440
            TabIndex        =   103
            Top             =   5040
            Width           =   1455
         End
         Begin VB.TextBox TxtName 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   3
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Text            =   "äÇã ˜ÇáÇ"
            Top             =   4680
            Width           =   1845
         End
         Begin VB.CheckBox ChkName 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ äÇã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   4440
            TabIndex        =   101
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox TxtRow 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   3
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Text            =   "ÑÏíÝ"
            Top             =   4200
            Width           =   1845
         End
         Begin VB.CheckBox ChkRow 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÑÏíÝ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   4440
            TabIndex        =   99
            Top             =   4080
            Width           =   1455
         End
         Begin VB.TextBox TxtFontGrid 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   3
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Text            =   "Titr"
            Top             =   2520
            Width           =   1845
         End
         Begin VB.TextBox TxtFontsizeGrid 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   3
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Text            =   "16"
            Top             =   3000
            Width           =   645
         End
         Begin VB.TextBox TxtHeaderFont 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   3
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Text            =   "Titr"
            Top             =   720
            Width           =   1845
         End
         Begin VB.TextBox TxtHeaderSize 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   3
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Text            =   "16"
            Top             =   1200
            Width           =   645
         End
         Begin VB.CheckBox ChkViewSegment 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÈÎÔ "
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   4560
            TabIndex        =   94
            Top             =   360
            Width           =   1455
         End
         Begin FLWCtrls.FWComboColor FWComboColorHeader 
            Height          =   315
            Index           =   3
            Left            =   600
            TabIndex        =   109
            Tag             =   "&H00000080&"
            Top             =   1680
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
         End
         Begin FLWCtrls.FWComboColor FWComboColorGrid 
            Height          =   315
            Index           =   3
            Left            =   600
            TabIndex        =   110
            Tag             =   "&H00000080&"
            Top             =   3600
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÑ ÊíÊÑ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   3000
            TabIndex        =   190
            Top             =   240
            Width           =   975
         End
         Begin VB.Line Line5 
            X1              =   0
            X2              =   6240
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2760
            TabIndex        =   122
            Top             =   5640
            Width           =   1095
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2760
            TabIndex        =   121
            Top             =   5160
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2760
            TabIndex        =   120
            Top             =   4680
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2760
            TabIndex        =   119
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "äæÚ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2760
            TabIndex        =   118
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÇíÒ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2760
            TabIndex        =   117
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Ñä Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2760
            TabIndex        =   116
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "ÌÏæá ˜ÇáÇåÇ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   3
            Left            =   4200
            TabIndex        =   115
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "äæÚ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2880
            TabIndex        =   114
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÇíÒ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2880
            TabIndex        =   113
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ñä Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2880
            TabIndex        =   112
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   3
            Left            =   4200
            TabIndex        =   111
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "ÈÎÔ 3"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   7215
         Index           =   2
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   840
         Width           =   6255
         Begin VB.TextBox TxtHeadeTitr 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Titr"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   187
            Text            =   "ÇÕáí 3"
            Top             =   240
            Width           =   1845
         End
         Begin VB.CheckBox ChkDescription 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÊæÖíÍÇÊ ˜ÇáÇ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   3840
            TabIndex        =   78
            Top             =   6480
            Width           =   2055
         End
         Begin VB.CheckBox ChkPicture 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ú˜Ó"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   4440
            TabIndex        =   77
            Top             =   6000
            Width           =   1455
         End
         Begin VB.TextBox TxtFee2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Text            =   "Ýí ÈíÑæä"
            Top             =   5640
            Width           =   1845
         End
         Begin VB.CheckBox ChkFee2 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ýí Ïæã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   4440
            TabIndex        =   75
            Top             =   5520
            Width           =   1455
         End
         Begin VB.TextBox TxtFee 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Text            =   "Ýí ÓÇáä"
            Top             =   5160
            Width           =   1845
         End
         Begin VB.CheckBox ChkFee1 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ýí Çæá"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   4440
            TabIndex        =   73
            Top             =   5040
            Width           =   1455
         End
         Begin VB.TextBox TxtName 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Text            =   "äÇã ˜ÇáÇ"
            Top             =   4680
            Width           =   1845
         End
         Begin VB.CheckBox ChkName 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ äÇã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   4440
            TabIndex        =   71
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox TxtRow 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Text            =   "ÑÏíÝ"
            Top             =   4200
            Width           =   1845
         End
         Begin VB.CheckBox ChkRow 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÑÏíÝ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   4440
            TabIndex        =   69
            Top             =   4080
            Width           =   1455
         End
         Begin VB.TextBox TxtFontGrid 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Text            =   "Titr"
            Top             =   2520
            Width           =   1845
         End
         Begin VB.TextBox TxtFontsizeGrid 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Text            =   "16"
            Top             =   3000
            Width           =   645
         End
         Begin VB.TextBox TxtHeaderFont 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Text            =   "Titr"
            Top             =   720
            Width           =   1845
         End
         Begin VB.TextBox TxtHeaderSize 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Text            =   "16"
            Top             =   1200
            Width           =   645
         End
         Begin VB.CheckBox ChkViewSegment 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÈÎÔ "
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   4560
            TabIndex        =   64
            Top             =   360
            Width           =   1455
         End
         Begin FLWCtrls.FWComboColor FWComboColorHeader 
            Height          =   315
            Index           =   2
            Left            =   600
            TabIndex        =   79
            Tag             =   "&H00000080&"
            Top             =   1680
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
         End
         Begin FLWCtrls.FWComboColor FWComboColorGrid 
            Height          =   315
            Index           =   2
            Left            =   600
            TabIndex        =   80
            Tag             =   "&H00000080&"
            Top             =   3600
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÑ ÊíÊÑ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   3000
            TabIndex        =   188
            Top             =   240
            Width           =   975
         End
         Begin VB.Line Line4 
            X1              =   0
            X2              =   6240
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2760
            TabIndex        =   92
            Top             =   5640
            Width           =   1095
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2760
            TabIndex        =   91
            Top             =   5160
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2760
            TabIndex        =   90
            Top             =   4680
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2760
            TabIndex        =   89
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "äæÚ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2760
            TabIndex        =   88
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÇíÒ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2760
            TabIndex        =   87
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Ñä Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2760
            TabIndex        =   86
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "ÌÏæá ˜ÇáÇåÇ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   2
            Left            =   4080
            TabIndex        =   85
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "äæÚ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2880
            TabIndex        =   84
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÇíÒ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2880
            TabIndex        =   83
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ñä Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2880
            TabIndex        =   82
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   2
            Left            =   4200
            TabIndex        =   81
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "ÈÎÔ 2"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   7215
         Index           =   1
         Left            =   -74280
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   840
         Width           =   6255
         Begin VB.TextBox TxtHeadeTitr 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Titr"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   185
            Text            =   "ÇÕáí 2"
            Top             =   240
            Width           =   1845
         End
         Begin VB.CheckBox ChkDescription 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÊæÖíÍÇÊ ˜ÇáÇ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   3840
            TabIndex        =   48
            Top             =   6480
            Width           =   2055
         End
         Begin VB.CheckBox ChkPicture 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ú˜Ó"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   4440
            TabIndex        =   47
            Top             =   6000
            Width           =   1455
         End
         Begin VB.TextBox TxtFee2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Text            =   "Ýí ÈíÑæä"
            Top             =   5640
            Width           =   1845
         End
         Begin VB.CheckBox ChkFee2 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ýí Ïæã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   4440
            TabIndex        =   45
            Top             =   5520
            Width           =   1455
         End
         Begin VB.TextBox TxtFee 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Text            =   "Ýí ÓÇáä"
            Top             =   5160
            Width           =   1845
         End
         Begin VB.CheckBox ChkFee1 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ýí Çæá"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   4440
            TabIndex        =   43
            Top             =   5040
            Width           =   1455
         End
         Begin VB.TextBox TxtName 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Text            =   "äÇã ˜ÇáÇ"
            Top             =   4680
            Width           =   1845
         End
         Begin VB.CheckBox ChkName 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ äÇã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   4440
            TabIndex        =   41
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox TxtRow 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Text            =   "ÑÏíÝ"
            Top             =   4200
            Width           =   1845
         End
         Begin VB.CheckBox ChkRow 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÑÏíÝ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   4440
            TabIndex        =   39
            Top             =   4080
            Width           =   1455
         End
         Begin VB.TextBox TxtFontGrid 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Text            =   "Titr"
            Top             =   2520
            Width           =   1845
         End
         Begin VB.TextBox TxtFontsizeGrid 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Text            =   "16"
            Top             =   3000
            Width           =   645
         End
         Begin VB.TextBox TxtHeaderFont 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Text            =   "Titr"
            Top             =   720
            Width           =   1845
         End
         Begin VB.TextBox TxtHeaderSize 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Text            =   "16"
            Top             =   1200
            Width           =   645
         End
         Begin VB.CheckBox ChkViewSegment 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÈÎÔ "
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   4560
            TabIndex        =   34
            Top             =   360
            Width           =   1455
         End
         Begin FLWCtrls.FWComboColor FWComboColorHeader 
            Height          =   315
            Index           =   1
            Left            =   600
            TabIndex        =   49
            Tag             =   "&H00000080&"
            Top             =   1680
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
         End
         Begin FLWCtrls.FWComboColor FWComboColorGrid 
            Height          =   315
            Index           =   1
            Left            =   600
            TabIndex        =   50
            Tag             =   "&H00000080&"
            Top             =   3600
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÑ ÊíÊÑ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   3000
            TabIndex        =   186
            Top             =   240
            Width           =   975
         End
         Begin VB.Line Line3 
            X1              =   0
            X2              =   6240
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2760
            TabIndex        =   62
            Top             =   5640
            Width           =   1095
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2760
            TabIndex        =   61
            Top             =   5160
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2760
            TabIndex        =   60
            Top             =   4680
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2760
            TabIndex        =   59
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "äæÚ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2760
            TabIndex        =   58
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÇíÒ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2760
            TabIndex        =   57
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Ñä Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2760
            TabIndex        =   56
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "ÌÏæá ˜ÇáÇåÇ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   1
            Left            =   4200
            TabIndex        =   55
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "äæÚ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2880
            TabIndex        =   54
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÇíÒ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2880
            TabIndex        =   53
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ñä Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2880
            TabIndex        =   52
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   1
            Left            =   4200
            TabIndex        =   51
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "ÈÎÔ 1"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   7215
         Index           =   0
         Left            =   -74280
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
         Width           =   6255
         Begin VB.TextBox TxtHeadeTitr 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Titr"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   183
            Text            =   "ÇÕáí 1"
            Top             =   240
            Width           =   1845
         End
         Begin VB.CheckBox ChkViewSegment 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÈÎÔ "
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   4560
            TabIndex        =   19
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox TxtHeaderSize 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Text            =   "16"
            Top             =   1200
            Width           =   645
         End
         Begin VB.TextBox TxtHeaderFont 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Text            =   "Titr"
            Top             =   720
            Width           =   1845
         End
         Begin VB.TextBox TxtFontsizeGrid 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Text            =   "16"
            Top             =   3000
            Width           =   645
         End
         Begin VB.TextBox TxtFontGrid 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Text            =   "Titr"
            Top             =   2520
            Width           =   1845
         End
         Begin VB.CheckBox ChkRow 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÑÏíÝ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   4440
            TabIndex        =   13
            Top             =   4080
            Width           =   1455
         End
         Begin VB.TextBox TxtRow 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Text            =   "ÑÏíÝ"
            Top             =   4200
            Width           =   1845
         End
         Begin VB.CheckBox ChkName 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ äÇã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   4440
            TabIndex        =   11
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox TxtName 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Text            =   "äÇã ˜ÇáÇ"
            Top             =   4680
            Width           =   1845
         End
         Begin VB.CheckBox ChkFee1 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ýí Çæá"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   4440
            TabIndex        =   9
            Top             =   5040
            Width           =   1455
         End
         Begin VB.TextBox TxtFee 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Text            =   "Ýí ÓÇáä"
            Top             =   5160
            Width           =   1845
         End
         Begin VB.CheckBox ChkFee2 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ýí Ïæã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   4440
            TabIndex        =   7
            Top             =   5520
            Width           =   1455
         End
         Begin VB.TextBox TxtFee2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Text            =   "Ýí ÈíÑæä"
            Top             =   5640
            Width           =   1845
         End
         Begin VB.CheckBox ChkPicture 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ Ú˜Ó"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   4440
            TabIndex        =   5
            Top             =   6000
            Width           =   1455
         End
         Begin VB.CheckBox ChkDescription 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÊæÖíÍÇÊ ˜ÇáÇ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   3840
            TabIndex        =   4
            Top             =   6480
            Width           =   2055
         End
         Begin FLWCtrls.FWComboColor FWComboColorHeader 
            Height          =   315
            Index           =   0
            Left            =   600
            TabIndex        =   18
            Tag             =   "&H00000080&"
            Top             =   1680
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
         End
         Begin FLWCtrls.FWComboColor FWComboColorGrid 
            Height          =   315
            Index           =   0
            Left            =   600
            TabIndex        =   20
            Tag             =   "&H00000080&"
            Top             =   3600
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÑ ÊíÊÑ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   3000
            TabIndex        =   184
            Top             =   240
            Width           =   975
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   6240
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   0
            Left            =   4200
            TabIndex        =   32
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ñä Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   2880
            TabIndex        =   31
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÇíÒ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   2880
            TabIndex        =   30
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "äæÚ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   2880
            TabIndex        =   29
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "ÌÏæá ˜ÇáÇåÇ"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   0
            Left            =   4080
            TabIndex        =   28
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Ñä Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   2760
            TabIndex        =   27
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "ÓÇíÒ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   2760
            TabIndex        =   26
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "äæÚ Þáã"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   2760
            TabIndex        =   25
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   2760
            TabIndex        =   24
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   2760
            TabIndex        =   23
            Top             =   4680
            Width           =   1095
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   2760
            TabIndex        =   22
            Top             =   5160
            Width           =   1095
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚäæÇä"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   2760
            TabIndex        =   21
            Top             =   5640
            Width           =   1095
         End
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmGoodMenuSetting.frx":A56A
      TabIndex        =   1
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ãÏíÑíÊ æ ÊäÙíãÇÊ äãÇíÔ ˜ÇáÇåÇ "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmGoodMenuSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Parameter() As Parameter
Dim i As Integer

Private Sub Form_Activate()
    VarActForm = Me.Name
    SetFirstToolBar
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
    
 '  Label10.Caption = "ãÏíÑíÊ æ ÊäÙíãÇÊ ÇíÓÊÇå ÔãÇÑå " & clsArya.StationNo
    
    formloadFlag = False
    Me.Left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 5000 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 0 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.Top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    If Me.Left < 0 Then Me.Left = 0
    If Me.Top < 0 Then Me.Top = 0
    If Me.Top > Me.ScaleHeight Then Me.Top = 0

    formloadFlag = True
    
    Dim i As Long
    
    For i = 0 To 5
    
        If clsGoodMenu.ViewSegmant(i) = True Then
           ChkViewSegment(i).Value = 1
        Else
           ChkViewSegment(i).Value = 0
        End If
        TxtHeadeTitr(i) = clsGoodMenu.HeaderTitr(i)
        TxtHeaderFont(i) = clsGoodMenu.HeaderFont(i)
        TxtHeaderSize(i) = clsGoodMenu.HeaderSizeFont(i)
        FWComboColorHeader(i).Color = Val(clsGoodMenu.HeaderColorFont(i))
        TxtFontGrid(i) = clsGoodMenu.GridFont(i)
        TxtFontsizeGrid(i) = clsGoodMenu.GridSizeFont(i)
        FWComboColorGrid(i).Color = Val(clsGoodMenu.GridColorFont(i))
        If clsGoodMenu.ViewRow(i) = True Then
           ChkRow(i).Value = 1
        Else
           ChkRow(i).Value = 0
        End If
        If clsGoodMenu.ViewName(i) = True Then
           ChkName(i).Value = 1
        Else
           ChkName(i).Value = 0
        End If
        If clsGoodMenu.ViewFee1(i) = True Then
           ChkFee1(i).Value = 1
        Else
           ChkFee1(i).Value = 0
        End If
        If clsGoodMenu.ViewFee2(i) = True Then
           ChkFee2(i).Value = 1
        Else
           ChkFee2(i).Value = 0
        End If
        If clsGoodMenu.ViewPicture(i) = True Then
           ChkPicture(i).Value = 1
        Else
           ChkPicture(i).Value = 0
        End If
        If clsGoodMenu.ViewDescription(i) = True Then
           ChkDescription(i).Value = 1
        Else
           ChkDescription(i).Value = 0
        End If
    
        TxtRow(i) = clsGoodMenu.RowName(i)
        TxtName(i) = clsGoodMenu.GoodName(i)
        TxtFee(i) = clsGoodMenu.Fee1Name(i)
        TxtFee2(i) = clsGoodMenu.Fee2Name(i)
    Next i
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

'    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    VarActForm = ""
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub

Public Sub Update()
    
    For i = 0 To 5
    
        clsGoodMenu.ViewSegmant(i) = ChkViewSegment(i).Value
        clsGoodMenu.HeaderTitr(i) = TxtHeadeTitr(i)
        clsGoodMenu.HeaderFont(i) = TxtHeaderFont(i)
        clsGoodMenu.HeaderSizeFont(i) = TxtHeaderSize(i)
        clsGoodMenu.HeaderColorFont(i) = FWComboColorHeader(i).Color
        clsGoodMenu.GridFont(i) = TxtFontGrid(i)
        clsGoodMenu.GridSizeFont(i) = TxtFontsizeGrid(i)
        clsGoodMenu.GridColorFont(i) = FWComboColorGrid(i).Color
        clsGoodMenu.ViewRow(i) = ChkRow(i).Value
        clsGoodMenu.ViewName(i) = ChkName(i).Value
        clsGoodMenu.ViewFee1(i) = ChkFee1(i).Value
        clsGoodMenu.ViewFee2(i) = ChkFee2(i).Value
        clsGoodMenu.ViewPicture(i) = ChkPicture(i).Value
        clsGoodMenu.ViewDescription(i) = ChkDescription(i).Value
        clsGoodMenu.RowName(i) = TxtRow(i)
        clsGoodMenu.GoodName(i) = TxtName(i)
        clsGoodMenu.Fee1Name(i) = TxtFee(i)
        clsGoodMenu.Fee2Name(i) = TxtFee2(i)
    Next i
  
    SetGoodMenuSettingFile
    
    Unload Me
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Public Sub SetFirstToolBar()
    
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(8).Enabled = True  'Enter
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True

End Sub

Public Sub ExitForm()

    Unload Me
    
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

