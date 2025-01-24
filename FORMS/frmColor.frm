VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmColor 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   9270
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Height          =   1695
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   840
      Width           =   2415
      Begin FLWCtrls.FWCheck FWSelFont 
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   64
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Value           =   0   'False
         Caption         =   "—‰ê “„Ì‰Â"
         Color           =   12582912
         ForeColor       =   128
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   12
         Alignment       =   1
      End
      Begin FLWCtrls.FWCheck FWSelFont 
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   65
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Value           =   0   'False
         Caption         =   "—‰ê ⁄‰«ÊÌ‰"
         Color           =   12582912
         ForeColor       =   128
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   12
         Alignment       =   1
      End
      Begin FLWCtrls.FWCheck FWSelFont 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   66
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Value           =   0   'False
         Caption         =   "—‰ê «ÿ·«⁄« "
         Color           =   12582912
         ForeColor       =   128
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   12
         Alignment       =   1
      End
      Begin FLWCtrls.FWCheck FWSelFont 
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   67
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Value           =   0   'False
         Caption         =   "—‰ê ÃœÊ·"
         Color           =   12582912
         ForeColor       =   128
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   12
         Alignment       =   1
      End
   End
   Begin FLWCtrls.FWRealButton FWBOK 
      Height          =   615
      Left            =   7080
      TabIndex        =   61
      Top             =   4920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Caption         =   " «∆Ìœ"
      ForeColor       =   16777152
      FontName        =   "B Homa"
      FontSize        =   12
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "‰„Ê‰Â"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2010
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   3840
      Width           =   1215
      Begin VB.CommandButton cmdPatern 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   600
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   " ‰ŸÌ„ œ” Ì —‰ê"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2025
      Index           =   0
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   3840
      Width           =   5160
      Begin VB.HScrollBar HScrBlue 
         Height          =   375
         Left            =   600
         Max             =   255
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   1440
         Width           =   3735
      End
      Begin VB.HScrollBar HScrGreen 
         Height          =   375
         Left            =   600
         Max             =   255
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   960
         Width           =   3735
      End
      Begin VB.HScrollBar HScrRed 
         Height          =   375
         Left            =   600
         Max             =   255
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label lblBlueReng 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   60
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label lblGreenReng 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   59
         Top             =   960
         Width           =   600
      End
      Begin VB.Label lblRedReng 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   58
         Top             =   480
         Width           =   600
      End
      Begin VB.Label lblBlueSel 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label lblGreenSel 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   960
         Width           =   480
      End
      Begin VB.Label lblRedSel 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3840
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   48
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2520
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   47
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   2040
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   46
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1560
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF00FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   45
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1080
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF80FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   44
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   600
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   43
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   120
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   42
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2520
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   41
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2040
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   40
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1560
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   39
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1080
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   38
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   600
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   37
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   2
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   600
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   3
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1080
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   4
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1560
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   5
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2040
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   6
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2520
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   7
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   8
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   600
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   9
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1080
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   10
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1560
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   11
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2040
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   12
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2520
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   13
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   14
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   600
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   15
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1080
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   16
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1560
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   17
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2040
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   18
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2520
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   19
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   20
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   600
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   21
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1080
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   22
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1560
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   23
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2040
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   24
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2520
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   25
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   26
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   27
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1080
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   28
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1560
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   29
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2040
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   30
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2520
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   31
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   32
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   33
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   34
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   35
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2040
         Width           =   450
      End
      Begin VB.OptionButton OptColor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   36
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2520
         Width           =   450
      End
   End
   Begin FLWCtrls.FWLabel FWLabel2 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "«‰ Œ«» —‰ê"
      FirstColor      =   16777152
      SecondColor     =   16744576
      Angle           =   0
      ForeColor       =   -2147483646
      FontName        =   "B Homa"
      FontSize        =   15.75
      Alignment       =   2
      Picture         =   "frmColor.frx":A4C2
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   5040
      OleObjectBlob   =   "frmColor.frx":A4DE
      TabIndex        =   62
      Top             =   720
      Width           =   480
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Obj As Object

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
Public Sub ExitForm()

    Unload Me
End Sub

Private Sub Form_Load()
If VarActForm = "frmInvoice" Or VarActForm = "frmPurchase" Then
   Me.Width = 9570
   FWLabel2.Width = 9650
Else
   Me.Width = 6885
   FWLabel2.Width = 6930
End If

'    formloadFlag = False
'    Me.Left = Val(GetSetting(strMainKey, Me.Name, "Left"))
'    If Val(GetSetting(strMainKey,Me.Name, "Height")) > 5000Then
'        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
'    End If
'    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 5000 Then
'        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
'    End If
'    Me.Top = Val(GetSetting(strMainKey, Me.Name, "Top"))
'    formloadFlag = True

'    m_oMySink.ShowSubSingleForm Me.Name
'    HandleStr = m_oMySink.GetMainWindow: SetParent Me.hwnd, HandleStr
    LoadForm Me.Name
    CenterCenterOffset Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Obj = Nothing

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
    
'    ShowMainMenu Me.Name


End Sub

Private Sub FWBOK_Click()
Dim varAnswer As Integer
ShowMessage " ¬Ì«  €ÌÌ—«  –ŒÌ—Â ‘Êœ ø", True, True, "»·Ì", "ŒÌ—"
varAnswer = modgl.mvarMsgIdx
If VarActForm <> "" Then
  If VarActForm = "frmInvoice" Then
    If varAnswer = vbYes Then
            If FWSelFont(0).Value = True Then
                Call SetUserSettingFile(cmdPatern.BackColor, 1)
                Call frmInvoice.ColorSetting
            End If
            If FWSelFont(1).Value = True Then
            End If
            If FWSelFont(2).Value = True Then
            End If
            If FWSelFont(3).Value = True Then
                Call SetUserSettingFile(cmdPatern.BackColor, 3)
                Call frmInvoice.ColorSetting
            End If
    End If
  ElseIf VarActForm = "frmPurchase" Then
    If varAnswer = vbYes Then
            If FWSelFont(0).Value = True Then
               Call SetUserSettingFile(cmdPatern.BackColor, 1)
               Call frmPurchase.ColorSetting
            End If
            If FWSelFont(1).Value = True Then
            End If
            If FWSelFont(2).Value = True Then
            End If
            If FWSelFont(3).Value = True Then
               Call SetUserSettingFile(cmdPatern.BackColor, 3)
               Call frmPurchase.ColorSetting
            End If
    End If
  End If
End If
  
If VarActForm = "" Then
   ShowMessage ".  ‰ŸÌ„«  «‰Ã«„ ‘œ", True, False, "ﬁ»Ê·", ""
   Unload Me
Else
   ShowMessage " ‰ŸÌ„«  «‰Ã«„ ‘œ.¬Ì«  €ÌÌ—«  œÌê—Ì Â„ œ«—Ìœø", True, True, "»·Ì", "ŒÌ—"
   varAnswer = modgl.mvarMsgIdx
   'VarAnswer = MsgBox(" ‰ŸÌ„«  «‰Ã«„ ‘œ.¬Ì«  €ÌÌ—«  œÌê—Ì Â„ œ«—Ìœø", vbYesNo)
   If varAnswer = vbNo Then
      Unload Me
   End If
End If
Exit Sub
Err1:
Resume Next
End Sub


Private Sub HScrBlue_Change()
   cmdPatern.BackColor = RGB(HScrRed.Value, HScrGreen.Value, HScrBlue.Value)
   lblBlueSel.BackColor = RGB(0, 0, HScrBlue.Value)
   lblBlueReng.Caption = HScrBlue.Value
End Sub

Private Sub HScrRed_Change()
    cmdPatern.BackColor = RGB(HScrRed.Value, HScrGreen.Value, HScrBlue.Value)
    lblRedSel.BackColor = RGB(HScrRed.Value, 0, 0)
    lblRedReng.Caption = HScrRed.Value
End Sub

Private Sub HScrGreen_Change()
   cmdPatern.BackColor = RGB(HScrRed.Value, HScrGreen.Value, HScrBlue.Value)
   lblGreenSel.BackColor = RGB(0, HScrGreen.Value, 0)
   lblGreenReng.Caption = HScrGreen.Value
End Sub

''Private Sub lstSayzKt_Click()
''    'cmdPatern.FontSize = Val(lstSayzKt.Text)
''    'pubbytFontSizeKlydKala = Val(lstSayzKt.Text)
''End Sub

Private Sub optColor_Click(index As Integer)
    cmdPatern.BackColor = OptColor(index).BackColor

'    publngBackColorKlydKala = optColor(Index).BackColor
End Sub



Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

