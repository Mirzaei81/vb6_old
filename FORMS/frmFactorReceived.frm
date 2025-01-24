VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmFactorReceived 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   12885
   Icon            =   "frmFactorReceived.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   5760
      Width           =   12735
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         ScaleHeight     =   435
         ScaleWidth      =   12195
         TabIndex        =   44
         Top             =   2040
         Width           =   12255
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Œ—ÊÃ « Ê„« Ìò Ê À»  Ê ç«Å  »⁄œ «“ œ—Ì«› "
            BeginProperty Font 
               Name            =   "B Nazanin"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   0
            Width           =   3975
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Œ—ÊÃ « Ê„« Ìò Ê À»  »⁄œ «“ œ—Ì«› "
            BeginProperty Font 
               Name            =   "B Nazanin"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   0
            Width           =   3495
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄œ„ Œ—ÊÃ « Ê„« Ìò"
            BeginProperty Font 
               Name            =   "B Nazanin"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "V1.0.0.2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   11400
            TabIndex        =   48
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "œ—Ì«›  ò«—  Ê ò«— -‰ﬁœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   120
         Width           =   6255
         Begin FLWCtrls.FWButton FWBtnOK 
            Height          =   735
            Left            =   2040
            TabIndex        =   40
            ToolTipText     =   "—òÊ—œÂ«Ì ‰„«Ì‘ œ«œÂ ‘œÂ œ— ÃœÊ· »«·« —«  À»  „Ì ò‰œ"
            Top             =   360
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   1296
            Caption         =   "À» (F12)"
            FontName        =   "B Homa"
            FontSize        =   11.25
            Alignment       =   1
            Object.ToolTipText     =   "—òÊ—œÂ«Ì ‰„«Ì‘ œ«œÂ ‘œÂ œ— ÃœÊ· »«·« —«  À»  „Ì ò‰œ"
         End
         Begin FLWCtrls.FWButton FWBtnPrint 
            Height          =   735
            Left            =   120
            TabIndex        =   41
            ToolTipText     =   "—òÊ—œÂ«Ì ‰„«Ì‘ œ«œÂ ‘œÂ œ— ÃœÊ· »«·« —«  ç«Å Ê À»  „Ì ò‰œ"
            Top             =   360
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   1296
            ButtonType      =   5
            Caption         =   "À»  Êç«Å(F6)"
            FontName        =   "B Homa"
            FontSize        =   11.25
            Alignment       =   1
            Object.ToolTipText     =   "—òÊ—œÂ«Ì ‰„«Ì‘ œ«œÂ ‘œÂ œ— ÃœÊ· »«·« —«  ç«Å Ê À»  „Ì ò‰œ"
         End
         Begin FLWCtrls.FWButton FWBtnPos 
            Height          =   735
            Left            =   3840
            TabIndex        =   42
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1296
            ButtonType      =   8
            Caption         =   "POs œ—Ì«›  «“ ÿ—Ìﬁ    (F11)"
            FontName        =   "B Homa"
            FontSize        =   9.75
            Alignment       =   1
         End
      End
      Begin VB.Frame Frame_Cash 
         Caption         =   "›ﬁÿ œ—Ì«›  ‰ﬁœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   120
         Width           =   4335
         Begin FLWCtrls.FWButton FWBtnCash_Print 
            Height          =   735
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1296
            ButtonType      =   7
            Caption         =   "œ—Ì«›  ‰ﬁœÌ Ê ç«Å(F8)"
            FontName        =   "B Homa"
            FontSize        =   9.75
            Alignment       =   1
         End
         Begin FLWCtrls.FWButton FWBtnCash 
            Height          =   735
            Left            =   2160
            TabIndex        =   38
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1296
            ButtonType      =   7
            Caption         =   "œ—Ì«›  ‰ﬁœÌ Ê À» (F7)"
            FontName        =   "B Homa"
            FontSize        =   9.75
            Alignment       =   1
         End
      End
      Begin FLWCtrls.FWButton FWBtnCancel 
         Height          =   615
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "«‰’—«› Ê Œ—ÊÃ"
         Top             =   720
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
         ButtonType      =   1
         Caption         =   "«‰’—«›"
         BackColor       =   12632256
         FontName        =   "B Homa"
         FontSize        =   11.25
         Alignment       =   1
         Object.ToolTipText     =   "«‰’—«› Ê Œ—ÊÃ"
      End
      Begin FLWCtrls.FWButton FWBtnCredit 
         Height          =   615
         Left            =   120
         TabIndex        =   49
         ToolTipText     =   "«‰ ﬁ«· »Â Õ”«» „‘ —Ì«‰ «⁄ »«—Ì"
         Top             =   120
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
         ButtonType      =   1
         Caption         =   "«⁄ »«—Ì Ê ç«Å"
         BackColor       =   12632256
         FontName        =   "B Homa"
         FontSize        =   9.75
         Alignment       =   1
         Object.ToolTipText     =   "«‰ ﬁ«· »Â Õ”«» „‘ —Ì«‰ «⁄ »«—Ì"
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   $"frmFactorReceived.frx":A4C2
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1680
         Width           =   12465
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   $"frmFactorReceived.frx":A55C
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1320
         Width           =   12465
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "„Õ«”»Â »«ﬁÌ„«‰œÂ ÊÃÊÂ œ—Ì«› Ì"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2655
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3000
      Width           =   12735
      Begin VB.CommandButton BtnKeypad 
         BackColor       =   &H8000000D&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Titr"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   0
         Left            =   3360
         TabIndex        =   25
         Tag             =   "0"
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   3
         Left            =   9891
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Tag             =   "3"
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   2
         Left            =   10824
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Tag             =   "2"
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   1
         Left            =   11760
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Tag             =   "1"
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   6
         Left            =   7092
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Tag             =   "6"
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   5
         Left            =   8025
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Tag             =   "5"
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   4
         Left            =   8958
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Tag             =   "4"
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   9
         Left            =   4293
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Tag             =   "9"
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   8
         Left            =   5226
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Tag             =   "8"
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   7
         Left            =   6159
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Tag             =   "7"
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "Å«ﬂ ﬂ‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   10
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1680
         Width           =   1875
      End
      Begin VB.TextBox TxtPayment 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8175
         TabIndex        =   8
         Top             =   960
         Width           =   1890
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "—Ì«·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   7440
         TabIndex        =   34
         Top             =   480
         Width           =   495
      End
      Begin VB.Label LblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label LblRemain 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   585
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»«ﬁÌ„«‰œÂ :"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ã„⁄ œ—Ì«› Ì : "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ ›«ﬂ Ê— :"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   465
         Left            =   10440
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   " œ—Ì«›  ÊÃÊÂ ‰ﬁœ  :  "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   10320
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label LblRemainPlus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   585
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   1905
      End
      Begin VB.Label LblRemainMinus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   585
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   1905
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "«÷«›Â"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   585
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "ò”—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "—Ì«·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   7440
         TabIndex        =   9
         Top             =   1080
         Width           =   495
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   240
      OleObjectBlob   =   "frmFactorReceived.frx":A5EE
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfgCheque 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   12735
      _cx             =   22463
      _cy             =   4048
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label LblUserName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9255
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   2505
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ«—»— :  "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   11760
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "œ—Ì«›  ›«ﬂ Ê—"
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
      Left            =   5385
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label LblPrePayment 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÅÌ‘ œ—Ì«› : "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "frmFactorReceived"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim Parameter() As Parameter
Dim Flag As Boolean
Public intSerialNo As Long
Dim Updated As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Dim bolSockIsConnectedPos As Boolean
Dim FinishedOk As Boolean
Dim PaymentRow As Long
Dim WithEvents ClsPos As AryaPosData.ClsTransaction
Attribute ClsPos.VB_VarHelpID = -1
Dim RPos As AryaPosData.PosResponse

Private Sub BtnKeypad_Click(index As Integer)
    If BtnKeypad(index).Tag = "" Then
        If Len(Trim(TxtPayment)) >= 1 Then
            TxtPayment = left(TxtPayment, Len(Trim(TxtPayment)) - 1)
        End If
    Else
        TxtPayment = TxtPayment + BtnKeypad(index).Tag
    End If

End Sub

Private Sub Form_Activate()
    If formloadFlag = False Then
        vsfgCheque.SetFocus
'        vsfgCheque.Select vsfgCheque.Row, vsfgCheque.Col: vsfgCheque.EditCell  'Sendkey "{F4}", False:
        formloadFlag = True
    End If
    TxtPayment_Change
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 27  ' Esc
                     FWBtnCancel_Click
                  Case vbKeyF7
                        FWBtnCash_Click
                  Case vbKeyF8
                        FWBtnCash_Print_Click
                  Case vbKeyF11
                  '' Temporary Disabled
'''                        FWBtnPos_Click
                  Case vbKeyF12  ' Esc
                        FWBtnOK_Click
                  Case vbKeyF6  ' Esc
                     If FWBtnPrint.Enabled = True Then
                        FWBtnPrint_Click
                     End If
              End Select
          Case 2
               Select Case KeyCode
                  Case 123  'Exit
                     If clsStation.KeyboardType = EnumKeyBoardType.Rb2 Then
                        FWBtnCancel_Click
                     End If
              End Select
    
    End Select
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrHandler
    Dim hMenu As Long

    hMenu = GetSystemMenu(Me.hWnd, False)

    DeleteMenu hMenu, 6, MF_BYPOSITION
    CenterTop Me
    
    SetGrid

    Label3.Caption = frmInvoice.lblSumPrice.Tag
    
    frmInvoice.MoveToCredit = False
    If frmInvoice.blnCreditCust = True Then FWBtnCredit.Enabled = True Else FWBtnCredit.Enabled = False
    FillGrid
  '  If vsfgCheque.Rows = vsfgCheque.FixedRows Then AddRowInGrid
    AddRowInGrid
'    If intSerialNo = 0 And mvarStatus <> Order Then
'        vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 1) = 5
'        vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 8) = frmInvoice.lblSumPrice.Tag - Val(frmInvoice.lblPayFactorTotal.Caption)
    If mvarStatus = Order Then
        vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 1) = 1
        vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 8) = frmInvoice.lblSumPrice.Tag - Val(frmInvoice.lblPayFactorTotal.Caption)
    End If
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@UserId", adInteger, 4, mvarCurUserNo)
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserName", Parameter)
    LblUserName.Caption = IIf(IsNull(Rst!AddUserName), "", Rst!AddUserName)
    Rst.Close
    
    LblPrePayment.Caption = frmInvoice.lblPayFactorTotal.Caption 'Because calculate in Rows
    LblTotal.Caption = DoCalculate
   ' LblRemain.Caption = CStr(Val(Label3.Caption) - Val(LblTotal.Caption) - Val(LblPrePayment.Caption))
    LblRemain.Caption = CStr(Val(Label3.Caption) - Val(LblTotal.Caption))

    Flag = True

    formloadFlag = False
    Me.left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 5000 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 5000 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    If Me.left < 0 Then Me.left = 0
    If Me.top < 0 Then Me.top = 0
    If Me.top > Me.ScaleHeight Then Me.top = 0

''    formloadFlag = True

'    WinsockListen.LocalPort = 24000
'    WinsockListen.Listen
    
    With vsfgCheque
        If .Rows > 1 Then
            PaymentRow = .Rows - 1
        '
        '        .ColHidden(3) = True
        '        .ColHidden(4) = True
        '        .ColHidden(6) = True
        '        .ColHidden(7) = True
        '        .ColHidden(8) = True
        '        .ColHidden(9) = True
        '        .ColHidden(10) = True
        '        .ColHidden(11) = True
        '        .ColHidden(12) = True
        '        .ColHidden(13) = True
        '
             If clsStation.PosPayment = True Then
                 .TextMatrix(PaymentRow, 1) = 5
                 InitPos
             Else
                 .TextMatrix(PaymentRow, 1) = 1
             End If
        End If
'
    End With
    
''''
''''    '''Temporary Disabled
''''    FWBtnPos.Enabled = False
''''''    If clsStation.PosPayment = True Then
''''''        If vsfgCheque.Cell(flexcpTextDisplay, PaymentRow, 11) = EnumPosType.PasargadPos Then   ' Pasargad
''''''            Set Pos = New PCPos
''''''            Pos.ComPort = Val(Right(vsfgCheque.Cell(flexcpTextDisplay, PaymentRow, 12), 1))
''''''        ElseIf vsfgCheque.Cell(flexcpTextDisplay, PaymentRow, 11) = EnumPosType.PersianPos Then
''''''
''''''            Set s = New PosInterface.SendInfo
''''''            s.InitializeLANSend vsfgCheque.Cell(flexcpTextDisplay, PaymentRow, 12), 17000
''''''
''''''        End If
''''''    Else
''''''        FWBtnPos.Enabled = False
''''''    End If
'''''    If clsArya.HardLockSerialNo = "93032703304" Then    'LoveOnSea
''''
''''        If clsStation.PosPayment = True Then
'''''            If vsfgCheque.Cell(flexcpTextDisplay, PaymentRow, 11) = EnumPosType.PasargadPos Then   ' Pasargad
'''''                Set Pos = New PCPos
'''''                Pos.ComPort = Val(Right(vsfgCheque.Cell(flexcpTextDisplay, PaymentRow, 12), 1))
''''            If vsfgCheque.Cell(flexcpTextDisplay, PaymentRow, 11) = EnumPosType.PersianPos Then
''''
''''                Set p = New PosInterface.PCPos
''''                Set h = frmFactorReceived
''''                p.InitLAN vsfgCheque.Cell(flexcpTextDisplay, PaymentRow, 12), 17000
'''''                Set p = New PosInterface.SendInfo
'''''                s.InitializeLANSend vsfgCheque.Cell(flexcpTextDisplay, PaymentRow, 12), 17000
''''                FWBtnPos.Enabled = True
''''            End If
''''        Else
''''            FWBtnPos.Enabled = False
''''        End If
'''''    End If

    If Val(GetSetting(strMainKey, Me.Name, "Option1")) = 0 Then
        Option1(0).Value = True
    ElseIf Val(GetSetting(strMainKey, Me.Name, "Option1")) = 1 Then
        Option1(1).Value = True
    ElseIf Val(GetSetting(strMainKey, Me.Name, "Option1")) = 2 Then
        Option1(2).Value = True
    End If


Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 3000

End Sub
Private Sub InitPos()
With vsfgCheque
    If clsStation.PosModel > 0 Then
        .TextMatrix(PaymentRow, 7) = clsStation.PosModel
        SetPosPort PaymentRow  ''«» œ« »«Ìœ ÅÊ“ «‰ Œ«» ‘Êœ
        .TextMatrix(PaymentRow, 11) = clsStation.PosModel
        .TextMatrix(PaymentRow, 12) = IIf(clsStation.PosPort = 0, clsStation.PosModel, clsStation.PosPort)    ''  PosPort
        .TextMatrix(PaymentRow, 13) = clsStation.PosModel
        .TextMatrix(PaymentRow, 14) = clsStation.PosModel
        .TextMatrix(PaymentRow, 15) = clsStation.PosModel
    Else
        .TextMatrix(PaymentRow, 7) = ""
        .TextMatrix(PaymentRow, 11) = ""
        .TextMatrix(PaymentRow, 12) = ""
        .TextMatrix(PaymentRow, 13) = ""
        .TextMatrix(PaymentRow, 14) = ""
        .TextMatrix(PaymentRow, 15) = ""
    End If
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top
    Set ClsPos = Nothing
    Set RPos = Nothing
End Sub

Private Sub FWBtnCancel_Click()
    If Flag = True Then
        frmMsg.fwlblMsg.Caption = "¬Ì« »—«Ì Œ—ÊÃ «“ ›—„ «ÿ„Ì‰«‰ œ«—Ìœ"
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = "»·Â"
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        frmMsg.Show vbModal
        If mvarMsgIdx <> vbYes Then Exit Sub
    End If
    mvarIndexNo = 0
    Unload Me
End Sub

Private Sub FWBtnCash_Click()
    
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 1) = 1
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 7) = ""
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 11) = ""
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 12) = ""
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 13) = ""
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 14) = ""
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 15) = ""
    If Update = True Then
        mvarIndexNo = 1
'        Updated = -1
'        Updated = frmInvoice.Update
'        If Updated > 0 Then
            Unload Me
'        End If
    Else
        ShowDisMessage "œ— À»  œ—Ì«›  „‘ò· ÊÃÊœ œ«—œ", 1000
    End If

End Sub

Private Sub FWBtnCash_Print_Click()
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 1) = 1
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 7) = ""
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 11) = ""
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 12) = ""
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 13) = ""
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 14) = ""
    vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 15) = ""
    If Update = True Then
        mvarIndexNo = 2
'        Updated = -1
'        Updated = frmInvoice.Update
'        If Updated > 0 Then
            Unload Me
'        End If
    Else
        ShowDisMessage "œ— À»  œ—Ì«›  „‘ò· ÊÃÊœ œ«—œ", 1000
    End If

End Sub

Private Sub FWBtnCredit_Click()
    frmInvoice.MoveToCredit = True
    FWBtnPrint_Click
End Sub

Private Sub FWBtnOK_Click()
    If Update = True Then
        mvarIndexNo = 1
'        Updated = -1
'        Updated = frmInvoice.Update
'        If Updated > 0 Then
            Unload Me
'        End If
    Else
        ShowDisMessage "œ— À»  œ—Ì«›  „‘ò· ÊÃÊœ œ«—œ", 1000
    End If
End Sub


Private Sub FWBtnPos_Click()

    If HasPcPos = False Then ShowDisMessage "«— »«ÿ « Ê„« Ìò »« ÅÊ“ »«‰òÌ »—ﬁ—«— ‰Ì”  .", 1500: Exit Sub
    
    On Error GoTo ErrHandler
    Dim Payment As Boolean
    Payment = False
    FinishedOk = False
With vsfgCheque
    For PaymentRow = 1 To vsfgCheque.Rows - 1
        If Val(vsfgCheque.TextMatrix(PaymentRow, 1)) = 5 Then
            Payment = True
            Exit For
        End If
    Next
    Dim strData As String
    strData = ""
    If Payment = False Then ShowDisMessage "‰ÕÊÂ Å—œ«Œ  «“ ÿ—Ìﬁ ò«—  »«‰òÌ ‰Ì” ", 1500: Exit Sub
    If Val(vsfgCheque.TextMatrix(PaymentRow, 8)) <= 0 Then MsgBox "„»·€ ’ÕÌÕ ‰Ì” ": Exit Sub
    
    ShowDisMessage "«—”«· „»·€ »Â ÅÊ“ »«‰òÌ , ·ÿ›« „‰ Ÿ— »„«‰Ìœ...", 1000
    
    FWBtnPos.Enabled = False

    Set ClsPos = New ClsTransaction
    Set RPos = New PosResponse
        
      '*****Config Database*****'
    '  ClsPos.SetPosDatabase "ServerName", "DbName", "DBLogin", "lemon7430"
    '  Set RPos = ClsPos.SetPos(3) ' Pos Id In Database
    '
    '
    '  '*****Config Manual*****'
    '  Set RPos = ClsPos.SetPosManual(POSType_PersianSwitch, ComunicationType_LAN, "127.0.0.1", 1250)
    Set RPos = ClsPos.SetPosManualInt(Val(vsfgCheque.Cell(flexcpTextDisplay, PaymentRow, 14)), Val(vsfgCheque.TextMatrix(PaymentRow, 12)), vsfgCheque.Cell(flexcpTextDisplay, PaymentRow, 13), vsfgCheque.Cell(flexcpTextDisplay, PaymentRow, 15))                '
 
    '*****Sale Function*****'
    ClsPos.Sale Val(frmInvoice.txtNo), vsfgCheque.ValueMatrix(PaymentRow, 8)


'    FWBtnCancel.Enabled = True
'    fComplete.DoSomething
    
'
'    End If
End With

Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 3000

End Sub

Private Sub ClsPos_DataRecieved(ByVal Result As AryaPosData.PosResponse)

    On Error GoTo ErrHandler
    Dim strData As String
  
    If Result.PRStatus Then
        
        strData = strData & "‰ ÌÃÂ :  " & Result.PRError & vbCrLf
        strData = strData & "„»·€ : " & Result.RPAmount & vbCrLf
        strData = strData & " —«ò‰‘ : " & Result.PRTranNumber & vbCrLf
        strData = strData & " “„«‰" & Mid(Result.RPDate, 9, 2) & ":" & Mid(Result.RPDate, 11, 2) & vbCrLf
        strData = strData & " : ò«—  " & Result.PRCardNumber & vbCrLf
        'strData = strData & " —„Ì‰«· : " & Result.PRTerminal & vbCrLf

        vsfgCheque.TextMatrix(PaymentRow, 9) = Result.PRTranNumber
        vsfgCheque.TextMatrix(PaymentRow, 10) = Result.PRCardNumber   'Card  Number
        FinishedOk = True
        FWBtnCancel.Enabled = False
        FWBtnCash.Enabled = False
        FWBtnCash_Print.Enabled = False

    Else
        FinishedOk = False
        strData = strData & " : ÅÌ€«„ Œÿ« " & Result.PRError & vbCrLf
        vsfgCheque.TextMatrix(PaymentRow, 1) = 1
        vsfgCheque.TextMatrix(PaymentRow, 7) = ""
        vsfgCheque.TextMatrix(PaymentRow, 11) = ""
        vsfgCheque.TextMatrix(PaymentRow, 12) = ""
        vsfgCheque.TextMatrix(PaymentRow, 13) = ""
        vsfgCheque.TextMatrix(PaymentRow, 14) = ""
        vsfgCheque.TextMatrix(PaymentRow, 15) = ""

    End If
    
    ShowMessage strData, True, False, "ﬁ»Ê·", ""
    
    If FinishedOk = False Then
        FWBtnPos.Enabled = True
    Else
        If Option1(0).Value = True Then
            ' No Operation
        ElseIf Option1(1).Value = True Then
            FWBtnOK_Click
        ElseIf Option1(2).Value = True Then
            FWBtnPrint_Click
        End If
    End If
    
Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 3000

End Sub

'

Private Sub Option1_Click(index As Integer)
    SaveSetting strMainKey, Me.Name, "Option1", index
End Sub

Private Sub FWBtnPrint_Click()
    If Update = True Then
        mvarIndexNo = 2
'        frmInvoice.Printing
        Unload Me
    Else
        ShowDisMessage "œ— À»  œ—Ì«›  „‘ò· ÊÃÊœ œ«—œ", 1000
    End If
End Sub
Private Function Update() As Boolean
    Update = False
    Dim i As Integer
    Dim st As String
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim c5 As String
    Dim c6 As String
    Dim c7 As String
    Dim c8 As String
    Dim c9 As String
    Dim c10 As String

    With vsfgCheque
        Dim FirstRowCash As Integer
        FirstRowCash = 0
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) = "1" Then 'Sum Of Cash In Flexgrid and will be one row
                If FirstRowCash = 0 Then
                    FirstRowCash = i
                Else
                    .TextMatrix(FirstRowCash, 8) = .TextMatrix(FirstRowCash, 8) + Val(.TextMatrix(i, 8))
                    .TextMatrix(i, 1) = ""
                    .TextMatrix(i, 8) = ""
                End If
            ElseIf .TextMatrix(i, 1) = "5" Then
                If .ValueMatrix(i, 7) = 0 Then
                    ShowDisMessage "‰Ê⁄ ÅÊ“»«‰òÌ —« „‘Œ’ ò‰Ìœ", 1000
                    Update = False
                    Exit Function
                End If
            End If
        Next i
            
    End With
   ' If Val(LblRemain.Caption) > 0 Then Exit Function
'    If Val(LblRemain.Caption) < 0 Then
'        frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ „»·€ «÷«›Ì «“ ’‰œÊﬁ Å—œ«Œ  ê—œœø"
'        frmMsg.fwBtn(0).ButtonType = flwButtonOk
'        frmMsg.fwBtn(0).Caption = "»·Â"
'        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
'        frmMsg.fwBtn(1).Caption = "ŒÌ—"
'        frmMsg.Show vbModal
'        If mvarMsgIdx <> vbYes Then Exit Function
'        If vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 1) <> "" And vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 8) <> "" Then
'            AddRowInGrid
'        End If
'        vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 1) = 1
'        vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, 8) = Val(LblRemain.Caption)
'        FWBtnOK.Enabled = True
'        FWBtnPrint.Enabled = True
'    End If
    st = ""
    Dim sumPrice As Double
    sumPrice = 0
    With vsfgCheque
        For i = 1 To .Rows - 1
            If Len(.TextMatrix(i, 1)) > 0 Then c1 = .TextMatrix(i, 1) Else c1 = "0"
            If Len(.TextMatrix(i, 2)) > 0 Then c2 = .TextMatrix(i, 2) Else c2 = "0"
            If Len(.TextMatrix(i, 3)) > 0 Then c3 = .TextMatrix(i, 3) Else c3 = "0"
            If Val(.TextMatrix(i, 4)) > 0# Then c4 = CStr(DateToNumber8(.TextMatrix(i, 4))) Else c4 = "0"
            If Len(.TextMatrix(i, 5)) > 0 Then c5 = .TextMatrix(i, 5) Else c5 = "0"
            If Len(.TextMatrix(i, 6)) > 0 Then c6 = .TextMatrix(i, 6) Else c6 = ""
            If Len(.TextMatrix(i, 7)) > 0 Then c7 = .TextMatrix(i, 7) Else c7 = "0"
            If Len(.TextMatrix(i, 8)) > 0 Then c8 = .TextMatrix(i, 8) Else c8 = "0"
            If Len(.TextMatrix(i, 9)) > 0 Then c9 = .TextMatrix(i, 9) Else c9 = ""
            If Len(.TextMatrix(i, 10)) > 0 Then c10 = .TextMatrix(i, 10) Else c10 = ""
            If mvarStatus = InvoiceReturn Then c8 = -1 * c8

            If Val(c1) > 0# And Abs(Val(c8)) > 0# Then st = GenerateDetailsStringFactorReceived(st, c1, c2, c3, c4, c5, c6, c7, c8, c9, c10)
            sumPrice = sumPrice + Abs(c8)
        Next i
    End With
    If sumPrice > 0 And sumPrice <> Val(Label3) And mvarStatus = Invoice Then
        ShowMessage "„»·€ Ê«—œ ‘œÂ »« „»·€ ›«ò Ê— Ìò”«‰ ‰Ì” ", True, False, "ﬁ»Ê·", ""
        Update = False
    Else
        sFactorReceived = st
        Flag = False
        Update = True
    End If
End Function

Private Function DoCalculate() As String
    Dim i As Integer
    Dim s As Double
    
    s = 0#
    For i = 1 To vsfgCheque.Rows - 1
        If Val(vsfgCheque.TextMatrix(i, 1)) > 0 And Val(vsfgCheque.TextMatrix(i, 1)) <> 4 Then
            s = s + Val(vsfgCheque.TextMatrix(i, 8))
        ElseIf Val(vsfgCheque.TextMatrix(i, 1)) = 4 Then
            s = s + Val(vsfgCheque.TextMatrix(i, 8)) '* Val(vsfgCheque.TextMatrix(i, 7))
        End If
    Next i
    DoCalculate = CStr(s)
End Function

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub TxtPayment_Change()
    If Val(LblRemain) - Val(TxtPayment.Text) > 0 Then
        LblRemainMinus.Caption = Val(LblRemain) - Val(TxtPayment.Text)
        LblRemainPlus.Caption = ""
        LblRemainMinus.Caption = Format(LblRemainMinus, "#,## —Ì«·")
    Else
        LblRemainPlus.Caption = Val(TxtPayment.Text) - Val(LblRemain)
        LblRemainMinus.Caption = ""
        LblRemainPlus.Caption = Format(LblRemainPlus, "#,## —Ì«·")
    End If

End Sub

Private Sub TxtPayment_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub vsfgCheque_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If formloadFlag = False Then Exit Sub
    With vsfgCheque
        If Not (.Col = 1 Or .Col = 7) Then Exit Sub
        If .ValueMatrix(Row, 1) = 5 And .ValueMatrix(Row, 7) > 0 Then
            SetPosPort Row
           ' .TextMatrix(Row, 7) = clsStation.PosModel
            .TextMatrix(Row, 11) = .TextMatrix(Row, 7)
            .TextMatrix(Row, 12) = IIf(clsStation.PosPort = 0, clsStation.PosModel, clsStation.PosPort)    ''  PosPort
            .TextMatrix(Row, 13) = .TextMatrix(Row, 7)
            .TextMatrix(Row, 14) = .TextMatrix(Row, 7)
            .TextMatrix(Row, 15) = .TextMatrix(Row, 7)
        Else
            .TextMatrix(Row, 7) = ""
            .TextMatrix(Row, 11) = ""
            .TextMatrix(Row, 12) = ""
            .TextMatrix(Row, 13) = ""
            .TextMatrix(Row, 14) = ""
            .TextMatrix(Row, 15) = ""
        End If
    End With
End Sub

Private Sub SetPosPort(Row As Long)
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblPub_Pos_ByStationId", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        Do While Not Rst.EOF
            If Rst.Fields("PosId").Value = Val(vsfgCheque.TextMatrix(Row, 7)) Then
                clsStation.PosPort = Rst.Fields("PortId").Value
                Exit Do
            End If
            Rst.MoveNext
        Loop
    End If
    If Rst.State <> 0 Then Rst.Close
    
End Sub

Private Sub vsfgCheque_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Integer

    For i = 1 To vsfgCheque.Rows - 1
        vsfgCheque.TextMatrix(i, 0) = CStr(i)
    Next i
End Sub

Private Sub vsfgCheque_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    For i = 0 To vsfgCheque.Cols - 1
        SaveSetting strMainKey, Me.Name, "Col" & Col, vsfgCheque.ColWidth(Col)
    Next

End Sub

Private Sub vsfgCheque_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case CInt(Val(vsfgCheque.TextMatrix(Row, 1)))
        Case 1
            If Col = 2 Or Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Or Col = 7 Then Cancel = True
        Case 2
            If Col = 7 Then Cancel = True
'        Case 3
'            If Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Or Col = 7 Or Col = 8 Then Cancel = True
        Case 3
            If Col = 2 Or Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Or Col = 7 Then Cancel = True
        Case 4
'            If Col = 2 Or Col = 3 Or Col = 5 Or Col = 6 Then Cancel = True
        Case 5
            If Col > 1 And Col < 7 Then Cancel = True
            'If Val(vsfgCheque.TextMatrix(Row, 7)) = 1# And FWBtnPrint.Visible = False Then Cancel = True
    End Select
End Sub

Private Sub SetGrid()
    With vsfgCheque
        .Rows = 2
        .Cols = 16
        .ColHidden(-1) = False
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "‰Ê⁄"
        .TextMatrix(0, 2) = "”—Ì«·"
        .TextMatrix(0, 3) = "‘„«—Â Õ”«»"
        .TextMatrix(0, 4) = " «—ÌŒ ”——”Ìœ"
        .TextMatrix(0, 5) = "»«‰ò"
        .TextMatrix(0, 6) = "‘⁄»Â"
        .TextMatrix(0, 7) = "ÅÊ“ »«‰òÌ"
        .TextMatrix(0, 8) = "„»·€"
        .TextMatrix(0, 9) = "‘„«—Â ÅÌ êÌ—Ì"
        .TextMatrix(0, 10) = "‘„«—Â ò«— "
        .TextMatrix(0, 11) = "‰Ê⁄ ÅÊ“"
        .TextMatrix(0, 12) = " ‰Ê⁄ ÅÊ— "
        .TextMatrix(0, 13) = "¬œ—” ÅÊ“"
        .TextMatrix(0, 14) = "‰Ê⁄ ÅÊ“"
        .TextMatrix(0, 15) = "”—Ì«· ÅÊ“"
        
        .ColHidden(2) = True
        .ColHidden(3) = True
        .ColHidden(4) = True
        .ColHidden(5) = True
        .ColHidden(6) = True

        .ColWidth(0) = 210
        .ColWidth(1) = 1200
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1200
        .ColWidth(6) = 2400
        .ColWidth(7) = 750
        .ColWidth(8) = 1200

        .FixedAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignCenterCenter
        .ColAlignment(10) = flexAlignCenterCenter
        .ColAlignment(11) = flexAlignCenterCenter
        .ColAlignment(12) = flexAlignCenterCenter
        .ColAlignment(13) = flexAlignCenterCenter
        .ColAlignment(14) = flexAlignCenterCenter
        .ColAlignment(15) = flexAlignCenterCenter

        .ColEditMask(4) = "##/##/##"
        .ColFormat(8) = "###,###"
   '     .ColHidden(14) = True
'        If clsArya.ExternalAccounting = False Then
'            Set Rst = RunStoredProcedure2RecordSet("Get_All_tRecvType")
'        Else
            Set Rst = RunStoredProcedure2RecordSet("Get_All_tRecvType_Acc")
'        End If
        .ColComboList(1) = .BuildComboList(Rst, "nvcDescription", "tintType")
        Rst.Close

        Set Rst = RunStoredProcedure2RecordSet("Get_All_tBanks")
        .ColComboList(5) = .BuildComboList(Rst, "nvcBankName", "tintBank")
        Rst.Close

        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
        Set Rst = RunParametricStoredProcedure2Rec("Get_tblPub_Pos_ByStationId", Parameter)
        .ColComboList(7) = .BuildComboList(Rst, "nvcBankName", "PosId")
        Rst.Close
        Set Rst = RunParametricStoredProcedure2Rec("Get_tblPub_Pos_ByStationId", Parameter)
        .ColComboList(11) = .BuildComboList(Rst, "PosName", "PosId")
        Rst.Close
        
        Set Rst = RunStoredProcedure2RecordSet("Get_All_PosPort")
        .ColComboList(12) = .BuildComboList(Rst, CStr("PortName"), "PortId")
        Rst.Close
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_tblPub_Pos_ByStationId", Parameter)
        .ColComboList(13) = .BuildComboList(Rst, "PosAddress", "PosId")
        Rst.Close
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_tblPub_Pos_ByStationId", Parameter)
        .ColComboList(14) = .BuildComboList(Rst, "nvcPosType", "PosId")
        Rst.Close
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_tblPub_Pos_ByStationId", Parameter)
        .ColComboList(15) = .BuildComboList(Rst, "nvcPosSerialNo", "PosId")
        Rst.Close
        
        Dim i As Long
        For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name, "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = .Width / 10     'Row
            End If
        Next i
    End With
End Sub

Private Sub FillGrid()
    Dim i As Integer
    ReDim Parameter(1 To 1) As Parameter

    Parameter(1) = GenerateInputParameter("@intSerialNo", adInteger, 4, intSerialNo)
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_FactorReceived_By_FactorSerial", Parameter)
    With vsfgCheque
        .Rows = .FixedRows
        Do While Rst.EOF = False
            If Rst!c8 < Label3 Then
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 1) = Rst!c1
            .TextMatrix(i, 2) = Rst!c2
            .TextMatrix(i, 3) = Rst!c3
            If Rst!c4 > 0 Then .TextMatrix(i, 4) = NumberToDate(Rst!c4)
            .TextMatrix(i, 5) = Rst!c5
            .TextMatrix(i, 6) = Rst!c6
            .TextMatrix(i, 7) = Rst!c7
            .TextMatrix(i, 8) = Rst!c8
            .TextMatrix(i, 9) = Rst!c9
            .TextMatrix(i, 10) = Rst!c10
            End If
            Rst.MoveNext
        Loop
    End With
    Rst.Close
End Sub

Private Sub AddRowInGrid()
    Dim flgAddRow As Boolean
    Dim s As String
    Dim C As Integer
    LblTotal.Caption = DoCalculate
    LblRemain.Caption = CStr(Val(Label3.Caption) - Val(LblTotal.Caption))
    If Val(LblRemain.Caption) < 0 Then LblRemain.Caption = 0
    If LblRemain.Caption = "0" Then Exit Sub
    flgAddRow = False
    If vsfgCheque.Rows > vsfgCheque.FixedRows Then
        s = ""
        For C = 1 To vsfgCheque.Cols - 1
            s = s + vsfgCheque.TextMatrix(vsfgCheque.Rows - 1, C)
        Next C
        If Len(s) > 0 Then flgAddRow = True
    Else
        flgAddRow = True
        ''Add Cash Record with zero price
'        vsfgCheque.Rows = vsfgCheque.Rows + 1
'        vsfgCheque.Row = vsfgCheque.Rows - 1
'        vsfgCheque.Col = 1
'        vsfgCheque.TextMatrix(vsfgCheque.Row, 8) = 0
    End If
    If flgAddRow = True Then
        
        vsfgCheque.Rows = vsfgCheque.Rows + 1
        vsfgCheque.Row = vsfgCheque.Rows - 1
        vsfgCheque.Col = 1
        vsfgCheque.TextMatrix(vsfgCheque.Row, 8) = LblRemain.Caption
    Else
        vsfgCheque.Row = vsfgCheque.Row + 1
        vsfgCheque.Col = 1
    End If
End Sub

Private Sub vsfgCheque_EnterCell()
With vsfgCheque
    .Select .Row, .Col: .EditCell
End With
End Sub

Private Sub vsfgCheque_KeyDown(KeyCode As Integer, Shift As Integer)
With vsfgCheque
If KeyCode = vbKeyReturn Then
    Select Case CInt(Val(vsfgCheque.TextMatrix(.Row, 1)))
        Case 1
            If .Col = 1 Then
                .Col = 8
            ElseIf .Col = 8 Then
                AddRowInGrid
            End If
        Case 2
            If .Col = 1 Or .Col = 2 Or .Col = 3 Or .Col = 4 Or .Col = 5 Then
                .Col = .Col + 1
            ElseIf .Col = 6 Then
                .Col = 8
            Else
                AddRowInGrid
            End If
        Case 3
            If .Col = 1 Then
                .Col = 8
'            ElseIf Col = 2 Then
'                If Val(vsfgCheque.TextMatrix(Row, 2)) = 0# Then
'                    frmMsg.fwlblMsg.Caption = "‘„«—Â »‰ —« Ê«—œ ‰„«ÌÌœ"
'                    frmMsg.fwBtn(0).Visible = False
'                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
'                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
'                    frmMsg.Show vbModal
'                    Exit Sub
'                End If
'
'                ReDim Parameter(1 To 1) As Parameter
'
'                Parameter(1) = GenerateInputParameter("@intCreditSerial", adInteger, 4, vsfgCheque.TextMatrix(Row, 2))
'                Set Rst = RunParametricStoredProcedure2Rec("Get_tFacCredit_IsUsed", Parameter)
'                If Rst!IsUsed > 0 Then
'                    frmMsg.fwlblMsg.Caption = "«“ «Ì‰ ‘„«—Â »‰ «” ›«œÂ ‘œÂ «” ° ‘„«—Â »‰ —« œÊ»«—Â Ê«—œ ‰„«ÌÌœ"
'                    frmMsg.fwBtn(0).Visible = False
'                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
'                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
'                    frmMsg.Show vbModal
'                    Rst.Close
'                    Exit Sub
'                End If
'                Rst.Close
'
'                ReDim Parameter(1) As Parameter
'
'                Parameter(0) = GenerateInputParameter("@intSerial", adInteger, 4, CLng(Val(vsfgCheque.TextMatrix(vsfgCheque.Row, 2))))
'                Parameter(1) = GenerateOutputParameter("@intAmount", adInteger, 4)
'                vsfgCheque.TextMatrix(vsfgCheque.Row, 8) = RunParametricStoredProcedure("Get_tCredit_Amount_BySerial", Parameter)
'                AddRowInGrid
            ElseIf .Col = 8 Then
                AddRowInGrid
            End If
        Case 4
        
        Case 5
            If .Col = 1 Then
                .Col = 7
                Sendkey "{F4}", False
'                .TextMatrix(.Row, 7) = "1"
            ElseIf .Col = 7 Then
                .Col = 8
            ElseIf .Col = 8 Then
                .Col = 9
            ElseIf .Col = 9 Then
                .Col = 10
            ElseIf .Col = 10 Then
                AddRowInGrid
            End If
    End Select
    LblTotal.Caption = DoCalculate
   ' LblRemain.Caption = CStr(Val(Label3.Caption) - Val(LblTotal.Caption) - Val(LblPrePayment.Caption))
    LblRemain.Caption = CStr(Val(Label3.Caption) - Val(LblTotal.Caption))
    If Val(LblRemain.Caption) < 0 Then
        FWBtnOK.Enabled = False
        FWBtnPrint.Enabled = False
    Else
        FWBtnOK.Enabled = True
        FWBtnPrint.Visible = True
        FWBtnPrint.Enabled = True
    End If
ElseIf KeyCode = vbKeyF6 Then
    If FWBtnPrint.Enabled = True Then
       FWBtnPrint_Click
    End If
ElseIf KeyCode = vbKeyF12 Then
    FWBtnOK_Click
ElseIf KeyCode = vbKeyF7 Then
      FWBtnCash_Click
ElseIf KeyCode = vbKeyF8 Then
      FWBtnCash_Print_Click
ElseIf KeyCode = vbKeyF11 Then
      FWBtnPos_Click
End If
End With
End Sub

Private Sub vsfgCheque_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Col
        Case 2, 3, 8
            If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn And KeyAscii <> 45 Then KeyAscii = 0
    End Select
End Sub

