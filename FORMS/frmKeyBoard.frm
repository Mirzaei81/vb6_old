VERSION 5.00
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmKeyBoard 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"frmKeyBoard.frx":0000
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15405
   Icon            =   "frmKeyBoard.frx":00BC
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   15405
   Begin VB.PictureBox PicKeyBoard 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H009F3A09&
      FillStyle       =   0  'Solid
      Height          =   2520
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   2460
      ScaleWidth      =   15345
      TabIndex        =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   15405
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   201
         Left            =   1380
         TabIndex        =   1
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "”"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   203
         Left            =   2655
         TabIndex        =   2
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "“"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   205
         Left            =   3885
         TabIndex        =   3
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "–"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   206
         Left            =   4530
         TabIndex        =   4
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "œ"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   207
         Left            =   5160
         TabIndex        =   5
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "Œ"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   208
         Left            =   5775
         TabIndex        =   6
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "Õ"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   209
         Left            =   6405
         TabIndex        =   7
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "ç"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   210
         Left            =   7020
         TabIndex        =   8
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "Ã"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   212
         Left            =   8295
         TabIndex        =   9
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   " "
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   213
         Left            =   8955
         TabIndex        =   10
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "Å"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   216
         Left            =   735
         TabIndex        =   11
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "Ì"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   217
         Left            =   1380
         TabIndex        =   12
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "Â"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   219
         Left            =   2655
         TabIndex        =   13
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "‰"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   220
         Left            =   3270
         TabIndex        =   14
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "„"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   221
         Left            =   3885
         TabIndex        =   15
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "·"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   222
         Left            =   4530
         TabIndex        =   16
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "ê"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   223
         Left            =   5160
         TabIndex        =   17
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "ﬂ"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   224
         Left            =   5775
         TabIndex        =   18
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "ﬁ"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   225
         Left            =   6390
         TabIndex        =   19
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "›"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   226
         Left            =   7020
         TabIndex        =   20
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "€"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   228
         Left            =   8295
         TabIndex        =   21
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "Ÿ"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   230
         Left            =   9630
         TabIndex        =   22
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "÷"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   231
         Left            =   10275
         TabIndex        =   23
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "’"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   263
         Left            =   720
         TabIndex        =   24
         Top             =   1620
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   979
         Caption         =   "·« Ì‰"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   264
         Left            =   4170
         TabIndex        =   25
         Top             =   1620
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   979
         Caption         =   "›«’·Â"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   265
         Left            =   8940
         TabIndex        =   26
         Top             =   1620
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   979
         Caption         =   "›«—”Ì"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   260
         Left            =   8295
         TabIndex        =   27
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "["
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   259
         Left            =   8955
         TabIndex        =   28
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "]"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   256
         Left            =   1380
         TabIndex        =   29
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "#"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   255
         Left            =   2010
         TabIndex        =   30
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "<"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   254
         Left            =   2655
         TabIndex        =   31
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   ">"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   248
         Left            =   5160
         TabIndex        =   32
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "("
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   247
         Left            =   5775
         TabIndex        =   33
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   ")"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   246
         Left            =   6405
         TabIndex        =   34
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "/"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   245
         Left            =   7020
         TabIndex        =   35
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "ø"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   253
         Left            =   4530
         TabIndex        =   36
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "∫"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   258
         Left            =   90
         TabIndex        =   37
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   ":"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   232
         Left            =   2655
         TabIndex        =   38
         Top             =   1620
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "¡"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   1650
         Index           =   266
         Left            =   60
         TabIndex        =   39
         Top             =   525
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   2910
         Caption         =   "»⁄œÌ "
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   244
         Left            =   12780
         TabIndex        =   40
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "1"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   243
         Left            =   13425
         TabIndex        =   41
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "2"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   242
         Left            =   14055
         TabIndex        =   42
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "3"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   241
         Left            =   12780
         TabIndex        =   43
         Top             =   555
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "4"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   240
         Left            =   13425
         TabIndex        =   44
         Top             =   555
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "5"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   239
         Left            =   14055
         TabIndex        =   45
         Top             =   555
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "6"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   238
         Left            =   12780
         TabIndex        =   46
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "7"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   237
         Left            =   13425
         TabIndex        =   47
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "8"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   236
         Left            =   14055
         TabIndex        =   48
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "9"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   234
         Left            =   12780
         TabIndex        =   49
         Top             =   1605
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   979
         Caption         =   "0"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   235
         Left            =   14055
         TabIndex        =   50
         Top             =   1620
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "."
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   267
         Left            =   3270
         TabIndex        =   51
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "{"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   202
         Left            =   2010
         TabIndex        =   52
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "é"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   251
         Left            =   14685
         TabIndex        =   53
         Top             =   525
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "+"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   252
         Left            =   14685
         TabIndex        =   54
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "-"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   249
         Left            =   14685
         TabIndex        =   55
         Top             =   1620
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "*"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   250
         Left            =   14685
         TabIndex        =   56
         Top             =   1065
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "/"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   214
         Left            =   9630
         TabIndex        =   57
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "»"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   229
         Left            =   8970
         TabIndex        =   58
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "ÿ"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   233
         Left            =   8310
         TabIndex        =   59
         Top             =   1620
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "¬"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   204
         Left            =   3270
         TabIndex        =   60
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "—"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   262
         Left            =   3885
         TabIndex        =   61
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "}"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   200
         Left            =   750
         TabIndex        =   62
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "‘"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   257
         Left            =   735
         TabIndex        =   63
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "!"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   211
         Left            =   7665
         TabIndex        =   64
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "À"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   227
         Left            =   7650
         TabIndex        =   65
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "⁄"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   261
         Left            =   7665
         TabIndex        =   66
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "°"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   585
         Index           =   268
         Left            =   9570
         TabIndex        =   67
         Top             =   0
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   1032
         Caption         =   "back"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   272
         Left            =   11490
         TabIndex        =   68
         Top             =   1320
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "7"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   18
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   271
         Left            =   11490
         TabIndex        =   69
         Top             =   420
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "8"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   18
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   215
         Left            =   10275
         TabIndex        =   70
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "«"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   269
         Left            =   11160
         TabIndex        =   71
         Top             =   870
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "<"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   18
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   585
         Index           =   270
         Left            =   11820
         TabIndex        =   72
         Top             =   840
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1032
         Caption         =   ">"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   18
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   273
         Left            =   7350
         TabIndex        =   73
         Top             =   1620
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   979
         Caption         =   "Õ–›"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   274
         Left            =   3300
         TabIndex        =   74
         Top             =   1620
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   979
         Caption         =   "Caps"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   275
         Left            =   10275
         TabIndex        =   75
         Top             =   1620
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "enter"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   218
         Left            =   2010
         TabIndex        =   76
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "Ê"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin FLWCtrls.FWRealButton FWKeyButton 
         Height          =   555
         Index           =   276
         Left            =   2010
         TabIndex        =   77
         Top             =   1620
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Caption         =   "∆"
         BackColor       =   16761087
         ForeColor       =   -2147483646
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         FontSize        =   15.75
      End
   End
End
Attribute VB_Name = "frmKeyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    OnTopMe Me, True
    PicKeyBoard.Visible = True
    CenterCenter Me
End Sub

Public Sub FWKeyButton_Click(Index As Integer)
KindKeys
KeyIndex = Index
If FWKeyButton(Index).Index = 274 Then
   KindKey = 2
End If
If FWKeyButton(Index).Index = 263 Then
   KindKey = 3
End If
If FWKeyButton(Index).Index = 265 Then
   KindKey = 1
End If
Select Case KindKey
       Case 1
            If FWKeyButton(Index).Index = 265 Then
               KindKey = 1
            End If
            Persian
       Case 2
            If FWKeyButton(Index).Index = 274 Then
               KindKey = 2
            End If
            CapsLock
       Case 3
            If FWKeyButton(Index).Index = 263 Then
               KindKey = 3
            End If
            Latin
End Select
End Sub

