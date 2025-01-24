VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmBenefit 
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16035
   Icon            =   "frmBenefit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   16035
   Begin VB.ListBox lstGoodLevel2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   10440
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   41
      Top             =   600
      Width           =   2625
   End
   Begin VB.ListBox lstGoodLevel1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   13200
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   40
      Top             =   600
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   31
      Top             =   480
      Width           =   5535
      Begin VB.ComboBox cmbSalMali 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton CmdUpdateHavaleh_Resid 
         Caption         =   "»Â —Ê“ —”«‰Ì ﬁÌ„   ÕÊ«·Â Ê —”Ìœ »—«Ì „Õ«”»Â ”Êœ ﬂ«·«Â«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton StoreDataUpdate 
         BackColor       =   &H00008080&
         Caption         =   "»Â —Ê“ —”«‰Ì ﬁÌ„   „«„ ‘œÂ ò«·«Â« »—«Ì „Õ«”»Â «—“‘ „ÊÃÊœÌ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "»—«Ì „Õ«”»Â «—“‘ „ÊÃÊœÌ ﬂ«·«Â« «” ›«œÂ „Ì ‘Êœ"
         Top             =   1080
         Width           =   2775
      End
      Begin FLWCtrls.FWProgressBar FWProgressBar1 
         Height          =   375
         Left            =   120
         Top             =   2040
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         BorderStyle     =   10
      End
      Begin MSMask.MaskEdBox txtDateTo 
         Height          =   465
         Left            =   3120
         TabIndex        =   35
         Top             =   1560
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   820
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtDateFrom 
         Height          =   465
         Left            =   3120
         TabIndex        =   36
         Top             =   1080
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   820
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«“  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”«· „«·Ì"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   5760
      TabIndex        =   26
      Top             =   480
      Width           =   4575
      Begin VB.CheckBox ChkFirstPrice 
         Alignment       =   1  'Right Justify
         Caption         =   "›ﬁÿ ò«·«Â«Ì »« ›Ì «Ê·ÌÂ ’›—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   57
         ToolTipText     =   "›ﬁÿ ﬂ«·«Â«Ì »« ›Ì «Ê·ÌÂ ’›— »« ¬Œ—Ì‰ ﬁÌ„  Œ—Ìœ »—Ê“ „Ì ‘Ê‰œ"
         Top             =   2115
         Width           =   2655
      End
      Begin VB.CommandButton cmdUpdateBuyPrice 
         Caption         =   "»Â —Ê“ —”«‰Ì ﬁÌ Œ—Ìœ »« ¬Œ—Ì‰ ﬁÌ„  Œ—Ìœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1770
         Width           =   1500
      End
      Begin VB.CommandButton cmdSetFirstPrice 
         Caption         =   "»Â —Ê“ —”«‰Ì ›Ì «Ê·ÌÂ »« ¬Œ—Ì‰ ﬁÌ„  Œ—Ìœ ﬁ»·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   105
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   945
         Width           =   1500
      End
      Begin VB.CommandButton cmdMojodiUpdate 
         BackColor       =   &H00008080&
         Caption         =   "»—Ê“ —”«‰Ì „ÊÃÊœÌ ﬂ«·«Â«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   48
         ToolTipText     =   "»—«Ì „Õ«”»Â «—“‘ „ÊÃÊœÌ ﬂ«·«Â« «” ›«œÂ „Ì ‘Êœ"
         Top             =   240
         Width           =   1485
      End
      Begin VB.Frame Frame28 
         Caption         =   "«‰»«—Â«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1440
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   120
         Width           =   2775
         Begin VB.ComboBox cmbBranch 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   360
            Width           =   2475
         End
         Begin VB.ComboBox cmbInventory 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   840
            Width           =   2475
         End
      End
      Begin VB.TextBox txtBarcode 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   1620
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "»«—òœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1680
         Width           =   585
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2895
      Left            =   120
      TabIndex        =   15
      Top             =   7800
      Width           =   5055
      Begin VB.Label LblTotalSellAmountLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "›—Ê‘ ﬂ· »« ò”—  Œ›Ì›« "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label LblTotalSellAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label LblTotalBenefitLossLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "”Êœ - “Ì«‰ ﬂ«·«Â«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label LblTotalBenefitLoss 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2280
         Width           =   2475
      End
      Begin VB.Label LblTotalSellReturnAmountLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "»—ê‘  «“ ›—Ê‘ ﬂ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label LblTotalSellReturnAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1200
         Width           =   2025
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "»Â«¡  „«„ ‘œÂ ›—Ê‘ ﬂ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label LblTotalFinalSellAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   720
         Width           =   2025
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "»Â«¡  „«„ ‘œÂ » «“ ›—Ê‘ ﬂ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label LblTotalFinalSellReturnAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1680
         Width           =   2025
      End
   End
   Begin VB.Frame Frame5 
      Height          =   2895
      Left            =   5280
      TabIndex        =   0
      Top             =   7800
      Width           =   10575
      Begin VB.Frame Frame_Acc 
         Caption         =   "”‰œ Õ”«»œ«—Ì"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   120
         Width           =   2175
         Begin VB.TextBox txtSanadNo 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   1320
            Width           =   1065
         End
         Begin VB.CommandButton cmdFirstMojodiSanad 
            BackColor       =   &H00008080&
            Caption         =   " Ê·Ìœ ”‰œ  „ÊÃÊœÌ «» œ«Ì œÊ—Â"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   52
            ToolTipText     =   "»—«Ì „Õ«”»Â «—“‘ „ÊÃÊœÌ ﬂ«·«Â« «” ›«œÂ „Ì ‘Êœ"
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "”‰œ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   1320
            Width           =   495
         End
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "«„ﬂ«‰ À»  „ÊÃÊœÌ «Ê·ÌÂ Ê ›Ì «Ê·ÌÂ Ê  Ê·Ìœ ”‰œ Õ”«»œ«—Ì »—«Ì „ÊÃÊœÌ «» œ«Ì œÊ—Â ÊÃÊœ œ«—œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Label LblTotalBuyReturnAmountLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "»—ê‘  «“ Œ—Ìœ ﬂ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1455
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "»Â«¡ ﬂ· «Ê·ÌÂ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label LblTotalBuyAmountLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ—Ìœ ﬂ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   975
         Width           =   1215
      End
      Begin VB.Label LblTotalMojodiPriceLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "«—“‘ ﬂ· „ÊÃÊœÌ «‰»«—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   2175
         Width           =   2415
      End
      Begin VB.Label LblTotalBuyReturnAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1455
         Width           =   2145
      End
      Begin VB.Label LblTotalLossAmountLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "÷«Ì⁄«  ﬂ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   255
         Width           =   1335
      End
      Begin VB.Label LblTotalLossAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label LblTotalFirstPrice 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   2025
      End
      Begin VB.Label LblTotalBuyAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   975
         Width           =   2025
      End
      Begin VB.Label LblTotalMojodiPrice 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   2175
         Width           =   2745
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "ÕÊ«·Â ﬂ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label LblTotalhavalehAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   2145
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "—”Ìœ ﬂ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label LblTotalResidAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1560
         Width           =   2025
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGood 
      Height          =   4680
      Left            =   120
      TabIndex        =   42
      Top             =   3120
      Width           =   15825
      _cx             =   27914
      _cy             =   8255
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      BackColorFixed  =   12648384
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   12648447
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   23
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBenefit.frx":A4C2
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   5
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
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   120
      Top             =   0
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   32896
      ForeColor2      =   128
      BackColor       =   9412754
      Caption         =   "„—Ê—"
      Alignment       =   2
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   360
      OleObjectBlob   =   "frmBenefit.frx":A6A5
      TabIndex        =   43
      Top             =   0
      Width           =   480
   End
   Begin FLWCtrls.FWToolTip FWToolTip 
      Left            =   0
      Top             =   0
      _ExtentX        =   926
      _ExtentY        =   926
      ForeColor       =   -2147483625
      BackColor       =   65535
   End
   Begin VB.Label lblGoodLevel2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "ê—ÊÂ ›—⁄Ì ò«·«Â«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   120
      Width           =   2025
   End
   Begin VB.Label lblGoodLevel1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ê—ÊÂ «’·Ì ò«·«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13440
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”Êœ Ê “Ì«‰ »«“—ê«‰Ì ﬂ«·«Â« Ê «—“‘ „ÊÃÊœÌ «Ê·ÌÂ"
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
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   -120
      Width           =   6015
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   6960
      TabIndex        =   44
      Top             =   4680
      Width           =   1215
   End
End
Attribute VB_Name = "frmBenefit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim tmpTextMatrix As String
Private Const IdxColGoodCode As Integer = 1
Dim Parameter() As Parameter
Dim TotalBenefitLoss As Currency
Dim TotalFirstPrice As Currency
Dim TotalBuyAmount As Currency
Dim TotalBuyReturnAmount As Currency
Dim TotalSellAmount As Currency
Dim TotalSellReturnAmount As Currency
Dim TotalMojodiPrice As Currency
Dim TotalHavalehAmount As Currency
Dim TotalResidAmount As Currency
Dim TotalLossAmount As Currency
Dim TotalFinalSellAmount As Currency
Dim TotalFinalSellReturnAmount As Currency
Dim GoodBenefitLoss As Currency
Dim SaleDiscountTotal As Currency
Dim i As Integer
Dim CrystalReport1
Dim CrystalReport2
    
Public Sub Find()
    frmFindGoods.Show vbModal
    
    i = vsGood.FindRow(mvarcode, 1, 1, True, True)
    If i > 0 Then
        vsGood.Row = i
        vsGood.ShowCell i, 0
    End If
End Sub

Public Sub ExitForm()
    Unload Me
End Sub

Public Sub SetFirstToolBar()

    Dim i As Integer
    
    AllButton vbOff, True
    mdifrm.Toolbar1.Buttons(7).Enabled = False   'Edit
    mdifrm.Toolbar1.Buttons(8).Enabled = False   'Apply
    mdifrm.Toolbar1.Buttons(9).Enabled = False   'Cancel
    mdifrm.Toolbar1.Buttons(13).Enabled = True   'Find
    
    mdifrm.Toolbar1.Buttons(15).Enabled = True  'printing
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    Select Case MyFormAddEditMode
    
        Case EnumAddEditMode.ViewMode
        
            For i = 7 To 9
                mdifrm.Toolbar1.Buttons(i).Enabled = True
            Next i
            vsGood.Editable = flexEDNone
            
        Case EnumAddEditMode.AddMode
        
            mdifrm.Toolbar1.Buttons(8).Enabled = True 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = True 'cancel key
 '           vsGood.Editable = flexEDKbdMouse
            
        Case EnumAddEditMode.EditMode
                    
            mdifrm.Toolbar1.Buttons(8).Enabled = True 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = True 'cancel key
'            vsGood.Editable = flexEDKbdMouse
    End Select
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
End Sub

Public Sub DefaultSetting()
    lstGoodLevel1.Clear
    lstGoodLevel2.Clear
    vsGood.Rows = 1
    
    If cmbInventory.ListIndex <> -1 Then
        FillLstGoodLevel1
    End If
End Sub

Public Sub FillLstGoodLevel1() ' it fills the lstGoodLevel1 using table tgoodlevel1
    On Error GoTo Err_Handler
    
    Dim Rst As New ADODB.Recordset
    
    lstGoodLevel1.Clear
    lstGoodLevel2.Clear
    vsGood.Rows = 1
    
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_Segment_Level1", Parameter)
        
    If (Rst.EOF = True And Rst.BOF = True) Then
        Exit Sub
    End If
    
    While Rst.EOF = False
        lstGoodLevel1.AddItem Rst.Fields("Description")
        lstGoodLevel1.ItemData(lstGoodLevel1.ListCount - 1) = Rst.Fields("Code")
        Rst.MoveNext
    Wend
    
    
    lstGoodLevel1.ListIndex = 0
    FillLstGoodLevel2
    Set Rst = Nothing
    
    Exit Sub
    
Err_Handler:
    LogSaveNew "frmbenefit => ", err.Description, err.Number, err.Source, "FillLstGoodLevel1"
    ShowErrorMessage
    err.Clear
End Sub

Public Sub FillLstGoodLevel2() ' it fills the lstGoodLevel2 using table tgoodlevel2
    On Error GoTo Err_Handler
    
    Dim Rst As New ADODB.Recordset
    Dim i As Integer
    Dim intSelectedItem As Integer
        
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    lstGoodLevel2.Clear
    vsGood.Rows = 1
    
    If lstGoodLevel1.ListIndex = -1 Then
        Set Rst = Nothing
        Exit Sub
    Else
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, lstGoodLevel1.ItemData(lstGoodLevel1.ListIndex))
        
        If Rst.State <> 0 Then Rst.Close
        Set Rst = RunParametricStoredProcedure2Rec("FillLstGoodLevel2", Parameter)
        If (Rst.EOF = True And Rst.BOF = True) Then
            Set Rst = Nothing
            Exit Sub
        End If
       ' rst.moveFirst
        While Rst.EOF = False
            Select Case clsStation.Language
                Case 0
                    lstGoodLevel2.AddItem Rst.Fields("Description")
                Case 1
                    lstGoodLevel2.AddItem Rst.Fields("LatinDescription")
            End Select
            
            lstGoodLevel2.ItemData(lstGoodLevel2.ListCount - 1) = Rst.Fields("Code")
            Rst.MoveNext
        Wend
        
        Set Rst = Nothing
        lstGoodLevel2.ListIndex = 0
        'FillvsGood
        CalculateTotalLabels
    End If
    
    Exit Sub
    
Err_Handler:
    LogSaveNew "frmBenefit => ", err.Description, err.Number, err.Source, "FillLstGoodLevel2"
    ShowErrorMessage
    err.Clear
End Sub

Public Sub BeforeUpdate()

End Sub

Public Sub Edit()
    With vsGood
        
 '       .Editable = flexEDKbdMouse

        MyFormAddEditMode = EnumAddEditMode.EditMode
        SetFirstToolBar
    End With
End Sub

Public Sub Update()
    
    Dim i As Long
    Dim j As Long
    Dim LongTemp As Integer
    Dim lngSelectedSubGroup  As Long
    
    Dim Rst As New ADODB.Recordset
    
    
    lngSelectedSubGroup = -1
    
    If MyFormAddEditMode = EnumAddEditMode.ViewMode Then Exit Sub
    
    vsGood_ValidateEdit vsGood.Row, vsGood.Col, False
    
    With vsGood
        If .Rows < 2 Then
            MyFormAddEditMode = EnumAddEditMode.ViewMode
            SetFirstToolBar
            Exit Sub
        End If
        
        For i = 1 To .Rows - 1
            .Row = i
            If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records
            
                
                
            End If
        Next i
        
        For j = 0 To lstGoodLevel2.ListCount - 1
            If lstGoodLevel2.Selected(j) = True Then
                lngSelectedSubGroup = j
                Exit For
            End If
        Next j

        Select Case MyFormAddEditMode
            Case EnumAddEditMode.EditMode
                For i = 1 To .Rows - 1
                    If InStr(.TextMatrix(i, 0), "*") > 0 Then 'Edited records
                        ReDim Parameter(4) As Parameter
                        Parameter(0) = GenerateInputParameter("@FirstMojodi", adDouble, 8, Val(.TextMatrix(i, 4)))
                        Parameter(1) = GenerateInputParameter("@FirstPrice", adDouble, 8, Val(.TextMatrix(i, 5)))
                        Parameter(2) = GenerateInputParameter("@Code", adInteger, 4, Val(Trim(.TextMatrix(i, 1))))
                        Parameter(3) = GenerateInputParameter("@IntInventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
                        Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
                        
                        RunParametricStoredProcedure "Update_Good_Store_FirstMojodi", Parameter
                    End If
                Next i
        End Select
        cmdMojodiUpdate_Click
'        CalculateTotalLabels
    End With
    
    Set Rst = Nothing
End Sub


Public Sub Cancel()
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    'FillvsGood
    CalculateTotalLabels
End Sub

Private Sub CheckFirstMojodi_Click()
    'FillvsGood
    CalculateTotalLabels
End Sub

Private Sub cmbInventory_Click()
   ' If cmbBranch.ListIndex = -1 Then Exit Sub
    FillLstGoodLevel1
    txtBarcode.SetFocus
    SetTooltipText
    mvarInventoryNo = cmbInventory.ItemData(cmbInventory.ListIndex)
End Sub

Private Sub FillSalMali()
    On Error GoTo Err_Handler
    
    cmbSalMali.Clear
    Set rs = RunStoredProcedure2RecordSet("Get_All_tAccountYears")
    Do While rs.EOF = False
        cmbSalMali.AddItem rs!AccountYear
        rs.MoveNext
    Loop
    Dim i As Integer
    For i = 0 To cmbSalMali.ListCount - 1
        cmbSalMali.ListIndex = i
        If AccountYear = cmbSalMali.Text Then
            Exit For
        End If
    Next
    'If cmbSalMali.ListCount > 0 Then cmbSalMali.ListIndex = 0
    rs.Close
    
    Exit Sub
    
Err_Handler:
    LogSaveNew "frmBenefit => ", err.Description, err.Number, err.Source, "FillSalMali"
    ShowErrorMessage
    err.Clear
End Sub

Private Sub cmdFirstMojodiSanad_Click()
    Dim ClosingSanadSaleGhabl As Integer
    Set rs = RunStoredProcedure2RecordSet("Get_All_tAccountYears")
    Do While rs.EOF = False
        If Val(cmbSalMali.Text) - 1 = rs!AccountYear Then
            ClosingSanadSaleGhabl = IIf(IsNull(rs!ClosingSanad), 0, rs!ClosingSanad)
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    If ClosingSanadSaleGhabl > 0 And Val(txtSanadNo) = 1 Then
        ShowDisMessage "”‰œ Õ”«»œ«—Ì „ÊÃÊœÌ «» œ«Ì œÊ—Â ﬁ»·« «“ ”«· „«·Ì ﬁ»·Ì »Â ”«· ÃœÌœ „‰ ﬁ· ‘œÂ", 2000
        Exit Sub
    End If
    If Val(TotalFirstPrice) <= 0 Then
        ShowDisMessage "«—“‘ „ÊÃÊœÌ «Ê·ÌÂ »—«Ì  Ê·Ìœ ”‰œ »«Ìœ »“—ê — «“ ’›— »«‘œ", 2000
        Exit Sub
    End If
    If Val(txtSanadNo) > 0 Then
        ShowMessage "”‰œ Õ”«»œ«—Ì „ÊÃÊœÌ «» œ«Ì œÊ—Â ﬁ»·«  Ê·Ìœ ‘œÂ. ¬Ì« „«Ì·Ìœ ¬‰ —« »« „ﬁ«œÌ— ÃœÌœ ÊÌ—«Ì‘ ﬂ‰Ìœø", True, True, "»·Ì", "ŒÌ—"
        If mvarMsgIdx = vbNo Then Exit Sub
    Else
        ShowMessage "œﬁ  ò‰Ìœ „ÊÃÊœÌ Â«Ì «Ê·ÌÂ «‰ ﬁ«·Ì «“ ”«· ﬁ»· œ«—«Ì „ﬁ«œÌ— Ê „»·€ ’ÕÌÕ »«‘‰œ. ¬Ì« „ÿ„∆‰ Â” Ìœø", True, True, "»·Ì", "ŒÌ—"
        If mvarMsgIdx = vbNo Then Exit Sub
    End If
    Dim SanadNo As Long
    If Val(TotalFirstPrice) > 0 Then
        SanadNo = Accounting.Insert_FirstMojodiDll(Val(txtSanadNo), Val(TotalFirstPrice))
        If SanadNo > 0 Then
           ShowDisMessage "”‰œ Õ”«»œ«—Ì »« ‘„«—Â  " & SanadNo & " »—«Ì À»  „ÊÃÊœÌ «» œ«Ì œÊ—Â «ÌÃ«œ Ì« ÊÌ—«Ì‘ ê—œÌœ", 1500
           ReDim Parameter(1) As Parameter
           Parameter(0) = GenerateInputParameter("@SanadNo", adInteger, 4, SanadNo)
           Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali))
           RunParametricStoredProcedure "Update_tAccountYears_FirstMojodi", Parameter
           txtSanadNo = SanadNo
        Else
           frmMsg.fwlblMsg.Caption = "œ—  Ê·Ìœ ”‰œ „ÊÃÊœÌ «» œ«Ì œÊ—Â „‘ﬂ· ÊÃÊœ œ«—œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
           frmMsg.fwBtn(0).ButtonType = flwButtonOk
           frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
           frmMsg.fwBtn(1).Visible = False
           frmMsg.Show vbModal
        End If
    End If

End Sub

Private Sub cmdMojodiUpdate_Click()
    If cmbInventory.ListIndex = -1 Then Exit Sub

    If Len(txtDateFrom.ClipText) < 6 Or Len(txtDateTo.ClipText) < 6 Then
        ShowDisMessage "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ ", 1000
        Exit Sub
    End If
    FWProgressBar1.Value = 0
    ReDim Parameter(11) As Parameter

    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    Parameter(6) = GenerateInputParameter("@Type", adInteger, 4, EnumGoodType.forBuySale)
    Parameter(7) = GenerateInputParameter("@InVentoryNo1", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(8) = GenerateInputParameter("@InVentoryNo2", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(9) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(10) = GenerateInputParameter("@UsePercentFlag", adInteger, 4, 0)
    Parameter(11) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    
    RunParametricStoredProcedure "Update_tblTotal_tInventory_tGood_For_Mojodi", Parameter
    FWProgressBar1.Value = 100

    ShowDisMessage " »Â —Ê“ —”«‰Ì «‰Ã«„ ‘œ ", 2000

    CalculateTotalLabels
End Sub

Private Sub cmdSetFirstPrice_Click()
    If cmbSalMali.Text = "" Or cmbInventory.ListIndex = -1 Then Exit Sub
    On Error GoTo ErrHandler
    
    Me.MousePointer = vbHourglass
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@AccountYear", adInteger, 4, Val(cmbSalMali.Text))
    If ChkFirstPrice.Value Then
        Parameter(1) = GenerateInputParameter("@Flag", adInteger, 4, 1) 'Goods that have zero for first price
    Else
        Parameter(1) = GenerateInputParameter("@Flag", adInteger, 4, 0)
    End If
    Parameter(2) = GenerateInputParameter("@InventoryNO", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    RunParametricStoredProcedure "Update_FirstPriceByBuyPrice", Parameter
    
    CalculateTotalLabels
    
    If ChkFirstPrice.Value Then
        frmDisMsg.lblMessage.Caption = " ﬂ«·«Â«Ì »« ›Ì «Ê·ÌÂ ’›— »Â —Ê“ —”«‰Ì ‘œ‰œ"
    Else
        frmDisMsg.lblMessage.Caption = " Â„Â ﬂ«·« »Â —Ê“ —”«‰Ì ‘œ‰œ"
    End If
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show
    Me.MousePointer = vbDefault
    
    CalculateTotalLabels
    
Exit Sub
ErrHandler:
    LogSaveNew "frmBenefit=>", err.Description, err.Number, err.Source, "CmdSetFirstPrice_Click"
    ShowErrorMessage
    Me.MousePointer = vbDefault

End Sub

Private Sub CmdUpdateHavaleh_Resid_Click()
    On Error GoTo Err_Handler
    If cmbInventory.ListIndex = -1 Then Exit Sub

    If Len(txtDateFrom.ClipText) < 6 Or Len(txtDateTo.ClipText) < 6 Then
        frmDisMsg.lblMessage = "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    If UpdateHavaleResid(cmbInventory.ItemData(cmbInventory.ListIndex), Val(cmbSalMali.Text), 0, txtDateFrom.Text, txtDateTo.Text) = True Then
        StoreDataUpdate_Click
        CalculateTotalLabels
    End If
    Exit Sub
    
Err_Handler:
    LogSaveNew "frmBenefit => ", err.Description, err.Number, err.Source, "CmdUpdateHavaleh_Resid_Click"
    ShowErrorMessage
    err.Clear
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdUpdateBuyPrice_Click()
    
    RunNonParametricStoredProcedure "Update_BuyPrice_by_LastPrice"
    ShowDisMessage "ﬁÌ„  Œ—Ìœ ò«·«Â« »« ¬Œ—Ì‰ ﬁÌ„  Œ—Ìœ ¬‰ ò«·« »Â —Ê“ ‘œ", 1500

End Sub

Private Sub Form_Activate()
    
    SetTooltipText
    VarActForm = Me.Name
    
    txtBarcode.Text = ""
   ' Frame3.BackColor = Me.BackColor
    Frame28.BackColor = Me.BackColor
    SetFirstToolBar
End Sub

Private Sub cmbSalMali_Click()
    txtDateFrom.Text = Right(cmbSalMali.Text, 2) & "/01" & "/01"
    If AccountYear = cmbSalMali.Text Then
        txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    Else
        If clsArya.MiladiDate = 0 Then
            txtDateTo.Text = Right(cmbSalMali.Text, 2) & "/12" & "/30"
        Else
            txtDateTo.Text = Right(cmbSalMali.Text, 2) & "/12" & "/31"
        End If
    End If
    'FillvsGood
    CalculateTotalLabels
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 13  ' Enter
                    SendKeys "{Left}", True
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
    On Error GoTo Err_Handler
    
    If ClsFormAccess.frmBenefit = False Then
        Unload Me
        Exit Sub
    End If
    
    If intVersion <> Diamond Then
        ShowDisMessage "«„ﬂ«‰ ê—› ‰ ”Êœ Ê “Ì«‰ ò«·«Â« ›ﬁÿ œ— ‰”ŒÂ «·„«” ÊÃÊœ œ«—œ", 1500
        Unload Me
        Exit Sub
    End If
    
    CenterTop Me
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

  
    With vsGood
        For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, "FrmBenefit_FlexGrid", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = .Width / 10     'Row
            End If
        Next i
        .Font.Name = "Tahoma"
        .Font.Size = 10
        .Font.Bold = False
    End With
    
    
    ChangeLanguage
    lstGoodLevel1.Clear
    lstGoodLevel2.Clear
    txtDateFrom.Text = Mid(AccountYear, 3, 2) & "/01" & "/01"
    txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    
    If clsArya.ExternalAccounting = True Then Frame_Acc.Enabled = True Else Frame_Acc.Enabled = False
    
    Set CrystalReport1 = CreateObject("Crystal.CrystalReport")
    Set CrystalReport2 = CreateObject("Crystal.CrystalReport")
    
    formloadFlag = True
    
    Exit Sub
    
Err_Handler:
    LogSaveNew "frmBenefit => ", err.Description, err.Number, err.Source, "Form_Load"
    ShowErrorMessage
    err.Clear
    Resume Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    Dim i As Integer
    
    AllButton vbOff, True
    
    Unload frmFindGoods
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
    Set CrystalReport1 = Nothing
    Set CrystalReport2 = Nothing
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub lstGoodLevel1_Click()
    FillLstGoodLevel2
End Sub

Private Sub lstGoodLevel1_ItemCheck(Item As Integer)
    
    Dim i As Integer
    
    If lstGoodLevel1.Selected(Item) = True Then
        For i = 0 To lstGoodLevel1.ListCount - 1
            If i <> Item And lstGoodLevel1.Selected(i) = True Then
                lstGoodLevel1.Selected(i) = False
            
            End If
        Next i
    End If
    
''''    FillvsGood
''''
''''    MyFormAddEditMode = EnumAddEditMode.ViewMode
''''    SetFirstToolbar
    
End Sub

Private Sub lstGoodLevel1_Scroll()
 '   FillLstGoodLevel2
End Sub

Private Sub lstGoodLevel2_ItemCheck(Item As Integer)
    'FillvsGood
    CalculateTotalLabels
End Sub

Public Sub ChangeLanguage()

    With vsGood
    
        .Cols = 21
        
        Select Case clsStation.Language
            Case Farsi
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "òœ"
                .TextMatrix(0, 2) = "‰«„ ò«·«"
                .TextMatrix(0, 3) = "»«—òœ"
                .TextMatrix(0, 4) = "„ÊÃÊœÌ «Ê·ÌÂ"
                .TextMatrix(0, 5) = "›Ì «Ê·ÌÂ"
                .TextMatrix(0, 6) = "«—“‘ »Â«Ì «Ê·ÌÂ"
                .TextMatrix(0, 7) = "Œ—Ìœ ﬂ·"
                .TextMatrix(0, 8) = "» «“ Œ—Ìœ"
                .TextMatrix(0, 9) = "«—“‘ ÷«Ì⁄« "
                .TextMatrix(0, 10) = "«—“‘ ÕÊ«·Â"
                .TextMatrix(0, 11) = "«—“‘ —”Ìœ"
                .TextMatrix(0, 12) = " „ÊÃÊœÌ"
                .TextMatrix(0, 13) = " ﬁÌ„   „«„ ‘œÂ"
                .TextMatrix(0, 14) = "«—“‘ „ÊÃÊœÌ"
                .TextMatrix(0, 15) = "›—Ê‘ ﬂ·"
                .TextMatrix(0, 16) = "»Â«¡ ›—Ê‘"
                .TextMatrix(0, 17) = "» «“ ›—Ê‘"
                .TextMatrix(0, 18) = "»Â«¡ » «“ ›—Ê‘"
                .TextMatrix(0, 19) = "”Êœ - “Ì«‰"
                .TextMatrix(0, 20) = "    "
                
            Case English
                .TextMatrix(0, 0) = "Row"
                .TextMatrix(0, 1) = "Code"
                .TextMatrix(0, 2) = "Name"
                .TextMatrix(0, 3) = "Barcode"
                .TextMatrix(0, 4) = "First Mojodi"
                .TextMatrix(0, 5) = "First Price"
                .TextMatrix(0, 6) = "Total First Price"
                .TextMatrix(0, 7) = "Total Buy Price"
                .TextMatrix(0, 8) = "Total Buy returned price"
                .TextMatrix(0, 9) = "Total Loss price"
                .TextMatrix(0, 10) = "Total From store price"
                .TextMatrix(0, 11) = "Total to store price"
                .TextMatrix(0, 12) = "Stock"
                .TextMatrix(0, 13) = "Stock Price"
                .TextMatrix(0, 14) = "Total current stock price"
                .TextMatrix(0, 15) = "Total sell price"
                .TextMatrix(0, 16) = "Sell Final Price"
                .TextMatrix(0, 17) = "Total sell returned price"
                .TextMatrix(0, 18) = "Return Sell Final price"
                .TextMatrix(0, 19) = "Total benefit - loss price"
                .TextMatrix(0, 20) = "    "
       End Select
        
'        .ColHidden(4) = True
'        .ColHidden(8) = True
'        .ColHidden(11) = True
        
'        .ColSort(10) = flexSortCustom
        .ColAlignment(-1) = flexAlignCenterCenter
'        .ColAlignment(25) = flexAlignRightCenter
        .FocusRect = flexFocusHeavy
'        .ColHidden(1) = True
'        .AutoSizeMode = flexAutoSizeColWidth
'        .AutoSize 0, .Cols - 2
'        .AutoSize 2, 14
        .AutoSearch = flexSearchFromCursor
    
        Dim i As Long
        For i = 3 To 19
            If i <> 4 And i <> 12 Then
                .ColFormat(i) = "###,###"
            End If
        Next
        .ColDataType(4) = flexDTLong
        .ColDataType(12) = flexDTLong
'        .ColFormat(4) = "###.000"
'        .ColFormat(12) = "###.000"
    End With
    
    FillBranch
    FillInventory
    FillSalMali
    DefaultSetting
            
    SetFirstToolBar

End Sub
Private Sub FillBranch()
    Dim i As Long
    Dim L_Rst As New ADODB.Recordset
    cmbBranch.Clear
    cmbBranch.AddItem "Â„Â ‘⁄»Â Â«"
    cmbBranch.ItemData(cmbBranch.NewIndex) = 0
    Set L_Rst = RunStoredProcedure2RecordSet("Get_All_Branches")
    
    Do While L_Rst.EOF = False
        cmbBranch.AddItem L_Rst!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = L_Rst!Branch
        L_Rst.MoveNext
    Loop
    
    L_Rst.Close: Set L_Rst = Nothing
    For i = 0 To cmbBranch.ListCount - 1
        If CurrentBranch = cmbBranch.ItemData(i) Then
            cmbBranch.ListIndex = i
            Exit For
        End If
    Next
End Sub
Private Sub FillInventory()
    cmbInventory.Clear
   ' If cmbBranch.ListIndex = -1 Then Exit Sub
    Dim rctmp As New ADODB.Recordset
    
    cmbInventory.Clear
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set rctmp = RunParametricStoredProcedure2Rec("GetInventory_Branch", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
            cmbInventory.AddItem rctmp.Fields("Description")
            cmbInventory.ItemData(cmbInventory.ListCount - 1) = Val(rctmp.Fields("InventoryNo"))
            rctmp.MoveNext
        Loop
        If rctmp.State = adStateOpen Then rctmp.Close
    End If
    
    Set rctmp = Nothing
  '  cmbInventory.ListIndex = 0

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub StoreDataUpdate_Click()
    If cmbInventory.ListIndex = -1 Then Exit Sub

    If Len(txtDateFrom.ClipText) < 6 Or Len(txtDateTo.ClipText) < 6 Then
        frmDisMsg.lblMessage = "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
       ' StoreDataUpdate.Enabled = False
        FWProgressBar1.Value = 0
        ReDim Parameter(8) As Parameter
    
        Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
        Parameter(3) = GenerateInputParameter("@DateBefore", adVarWChar, 50, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(4) = GenerateInputParameter("@DateAfter", adVarWChar, 50, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(5) = GenerateInputParameter("@Type", adInteger, 4, EnumGoodType.forBuySale)
        Parameter(6) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
        Parameter(7) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        Parameter(8) = GenerateInputParameter("@ZeroNegative", adBoolean, 1, 1)
        
        RunParametricStoredProcedure "Update_tblTotal_tInventory_tGood_For_FinalPrice", Parameter
        FWProgressBar1.Value = 100

'        Set Rst = RunParametricStoredProcedure2Rec("GetInventoryAtomicReport_Mojodi", Parameter )
    
        DefaultSetting
        FWProgressBar1.Value = 0
        StoreDataUpdate.Enabled = True
        frmDisMsg.lblMessage = " »Â —Ê“ —”«‰Ì ﬁÌ„   „«„ ‘œÂ ò«·«Â««‰Ã«„ ‘œ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal

End Sub

Private Sub txtBarcode_Change()
    If Right(txtBarcode.Text, 1) = "/" Then
        txtBarcode.Text = Left(txtBarcode.Text, Len(txtBarcode.Text) - 1)
    ElseIf Left(txtBarcode.Text, 1) = "/" Then
        txtBarcode.Text = Right(txtBarcode.Text, Len(txtBarcode.Text) - 1)
    End If
    If Len(txtBarcode.Text) > 2 Then
    If Asc(Mid(txtBarcode.Text, Len(txtBarcode.Text) - 1, 1)) = 13 Then
        txtBarcode.Text = Left(txtBarcode.Text, Len(txtBarcode.Text) - 2)
    End If
    End If
    i = vsGood.FindRow(Trim(txtBarcode.Text), 1, 3, True, True)
    If i > 0 Then
        vsGood.Row = i
        vsGood.ShowCell i, 3
    Else
        vsGood.Row = 0
        vsGood.ShowCell 0, 0
    End If

End Sub

Private Sub txtBarcode_GotFocus()
    SetKbLayout LANG_EN_US
    txtBarcode.Text = ""
End Sub
Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
        Case 0
            Select Case KeyCode
                Case 13
                    vsGood.SetFocus
                   ' KeyCode = 0
                 '   txtBarcode.Text = ""
                    If i > 0 Then
                        vsGood.Row = i
                        vsGood.ShowCell i, 3
                        vsGood.Row = i
                        vsGood.Col = 5
               '         vsGood.Selec vsGood.Row, vsGood.Col
                        vsGood.EditCell
                    End If
            End Select
    End Select
End Sub

Private Sub vsGood_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsGood
        If (.TextMatrix(Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And Col > 1 And tmpTextMatrix <> .TextMatrix(Row, Col) Then
        
            If MyFormAddEditMode = EnumAddEditMode.EditMode And InStr(.TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = Trim(.TextMatrix(Row, 0)) & "*"
            End If
            
        Else

        End If
      '  .AutoSizeMode = flexAutoSizeColWidth
      '  .AutoSize Col, Col
        

    End With


End Sub

Private Sub vsGood_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    For i = 0 To vsGood.Cols - 1
        SaveSetting strMainKey, "FrmBenefit_FlexGrid", "Col" & i, vsGood.ColWidth(i)
    Next
End Sub

'Private Sub vsGood_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    tmpTextMatrix = vsGood.TextMatrix(Row, Col)
'End Sub

Private Sub vsGood_BeforeSort(ByVal Col As Long, Order As Integer)
If Col = 6 Or Col = 14 Or Col = 19 Then
    With vsGood
        For i = 1 To .Rows - 1
            If InStr(1, .TextMatrix(i, Col), "-", vbTextCompare) Then
                .TextMatrix(i, Col) = -1 * (.TextMatrix(i, Col))
                .TextMatrix(i, Col) = -1 * (.TextMatrix(i, Col))
            End If
        Next i
    End With
End If
End Sub
Private Sub vsGood_AfterSort(ByVal Col As Long, Order As Integer)
If Col = 6 Or Col = 14 Or Col = 19 Then
    With vsGood
        For i = 1 To .Rows - 1
            If InStr(1, .TextMatrix(i, Col), "-", vbTextCompare) Then
                .TextMatrix(i, Col) = -1 * (.TextMatrix(i, Col))
                .TextMatrix(i, Col) = (.TextMatrix(i, Col)) & "-"
            End If
        Next i
    End With
End If
End Sub

Private Sub vsGood_Click()
    With vsGood
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And (.Col = 4 Or .Col = 5) Then
            If (.Col = 4 Or .Col = 5) And ClsFormAccess.FirstMojodiControl = True Then
               .Select .Row, .Col
               .EditCell
            Else
                ShowDisMessage "‘„« «Ã«“Â œ” —”Ì »Â «Ì‰ ﬁ”„  —« ‰œ«—Ìœ", 2000
            End If
        Else
            If .Col = 2 And ClsFormAccess.frmGoodTurnOver = True Then
                frmGoodTurnOver.Show
                frmGoodTurnOver.cmbInventory.ListIndex = cmbInventory.ListIndex
                frmGoodTurnOver.cmbSalMali.ListIndex = cmbSalMali.ListIndex
                
                frmGoodTurnOver.fwBtnGoodFind.Caption = .TextMatrix(.Row, 2)
                frmGoodTurnOver.fwBtnGoodFind.Tag = .TextMatrix(.Row, 1)
                frmGoodTurnOver.txtDateFrom.Text = txtDateFrom.Text
                frmGoodTurnOver.txtDateTo.Text = txtDateTo.Text
                frmGoodTurnOver.StoreDataUpdate.Enabled = True
                frmGoodTurnOver.txtBarcode.Text = .TextMatrix(.Row, 3)
                frmGoodTurnOver.StoreDataUpdate_Click
            End If
        End If
        
    End With

End Sub

Private Sub vsGood_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Then Exit Sub
    With vsGood
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And (.Col = 4 Or .Col > 5) Then
            If (.Col = 4 Or .Col = 5) And ClsFormAccess.FirstMojodiControl = True Then
               .Select .Row, .Col
               .EditCell
            Else
                ShowDisMessage "‘„« «Ã«“Â œ” —”Ì »Â «Ì‰ ﬁ”„  —« ‰œ«—Ìœ", 2000
            End If
        End If
    End With
    
End Sub


Private Sub vsGood_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsGood
        If KeyAscii = 39 Then KeyAscii = 0
        
        If ((Col <> 4 And Col <> 5)) Or (IsNumeric(Chr(KeyAscii)) = False And KeyAscii = 8) Then
            
            KeyAscii = 0
            
        ElseIf IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 Then
            
            KeyAscii = 0
            
        ElseIf (Col <> 4 And Col <> 5) Or KeyAscii = 8 Then
            
            KeyAscii = 0
            
        ElseIf MyFormAddEditMode = EditMode Then
            
            If Row > 0 And InStr(1, .TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = .TextMatrix(Row, 0) & "*"
            End If
            
        End If
        
    End With
    
End Sub


Private Sub vsGood_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGood
        .Row = Row
        .Col = Col
    End With
End Sub

Public Sub Printing()
    
    On Error GoTo Err_Handler
    If lstGoodLevel1.ListCount < 1 Then Exit Sub
    If lstGoodLevel2.ListCount < 1 Then Exit Sub
    
    Dim i, j As Long
    Dim intSelectedLevel1 As Integer
    intSelectedLevel1 = -1
    
    For i = 0 To lstGoodLevel1.ListCount - 1
        If lstGoodLevel1.Selected(i) = True Then
            intSelectedLevel1 = i
            Exit For
        End If
    Next i
        
    Dim level1 As Integer
    If intSelectedLevel1 <> -1 Then
        level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
    Else
        level1 = -1
    End If
    
    frmInput.fwlblInput.Caption = "‰Ê⁄ ê“«—‘ "
'    frmInput.OptionLevel(0).Caption = "ò«·«Â«Ì »« ê—œ‘"
'    frmInput.OptionLevel(1).Caption = "ò·ÌÂ ò«·«Â«"
'    frmInput.OptionLevel(2).Caption = " ò«·«Â«Ì »« ê—œ‘ Ê „ÊÃÊœÌ œ«—"
    frmInput.OptionLevel(0).Caption = "ê“«—‘ „ÊÃÊœÌ ò«·«Â«"
    frmInput.OptionLevel(1).Caption = "ê“«—‘ ”Êœ Ê “Ì«‰ ò«·«Â«"
    frmInput.OptionLevel(0).Value = True
    frmInput.btnCancel.Visible = True
    frmInput.Picture1.Visible = True
    frmInput.txtInput.Visible = False
    frmInput.OptionLevel(2).Visible = False
    frmInput.Show vbModal
    If mvarInput = "" Then
        Exit Sub
    End If
    CrystalReport1.ReportFileName = ""
    Dim fileSystem As New FileSystemObject
    Dim IsFileExist As Boolean
    Dim intIndex As Integer
    If mvarInput = "0" Then
        ReDim Parameter(9) As Parameter
        
        Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, clsDate.shamsi(Date))
        Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
        Parameter(3) = GenerateInputParameter("@Date1", adVarWChar, 50, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(4) = GenerateInputParameter("@Date2", adVarWChar, 50, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(5) = GenerateInputParameter("@AccountYear1", adSmallInt, 2, Val(cmbSalMali.Text))
        Parameter(6) = GenerateInputParameter("@Inventory1", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
'        If mvarInput = "0" Then
'            Parameter(7) = GenerateInputParameter("@Flag1", adInteger, 4, 1)
'        ElseIf mvarInput = "1" Then
            Parameter(7) = GenerateInputParameter("@Flag1", adInteger, 4, 2)  ''''all goods
'        ElseIf mvarInput = "2" Then
'            Parameter(7) = GenerateInputParameter("@Flag1", adInteger, 4, 3)
'        End If
        Parameter(8) = GenerateInputParameter("@Level11", adInteger, 4, level1)
        Parameter(9) = GenerateInputParameter("@Level12", adInteger, 4, level1)
        
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepInventoryGood_Mojodi_All_A4.rpt"
        CrystalReport1.ReportTitle = "  ê“«—‘ „ÊÃÊœÌ Ã‰”Ì «‰»«—-" & cmbInventory.Text
    
        IsFileExist = fileSystem.FileExists(CrystalReport1.ReportFileName)
        If IsFileExist = False Then
            frmDisMsg.lblMessage = " ›«Ì· ê“«—‘ entoryGood_Mojodi_All_A4.rpt ÅÌœ« ‰‘œ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            Exit Sub
        End If
        For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
            CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
        Next intIndex
        
'        For intIndex = UBound(Parameter) - LBound(Parameter) + 1 To 30
'             CrystalReport1.ParameterFields(intIndex) = ""
'        Next intIndex
        
        CrystalReport1.WindowShowGroupTree = True
        CrystalReport1.WindowShowSearchBtn = True
        CrystalReport1.WindowState = crptMaximized
        ODBCSetting clsArya.ServerName, clsArya.DbName
        CrystalReport1.Connect = CrystallConnection
        CrystalReport1.Action = 1
        CrystalReport1.RetrieveDataFiles
        
        If Screen.Width > 12000 Then
            CrystalReport1.PageZoom (100)
        Else
            CrystalReport1.PageZoom (75)
        End If
    
    ElseIf mvarInput = "1" Then
        
        ReDim Parameter(5)
        Parameter(0) = GenerateInputParameter("@DateBefore", adVarWChar, 8, Me.txtDateFrom.Text)
        Parameter(1) = GenerateInputParameter("@DateAfter", adVarWChar, 8, Me.txtDateTo.Text)
        Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        Parameter(3) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
        Parameter(4) = GenerateInputParameter("@GoodLevel1", adInteger, 4, level1)
        Parameter(5) = GenerateInputParameter("@SelectedLevelsString", adVarWChar, 4000, strSelectedLevels)
        
'        Set L_Rst = RunParametricStoredProcedure2Rec("Get_Benefit_Loss", Parameter)
        CrystalReport2.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSoodZianBazargani.rpt"
        CrystalReport2.ReportTitle = "  ê“«—‘ ”Êœ Ê “Ì«‰ «‰»«— -" & cmbInventory.Text
        CrystalReport2.Destination = crptToWindow 'crptToPrinter '
    
        IsFileExist = fileSystem.FileExists(CrystalReport2.ReportFileName)
        If IsFileExist = False Then
            frmDisMsg.lblMessage = " ›«Ì· ê“«—‘ RepSoodZianBazargani.rpt ÅÌœ« ‰‘œ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            Exit Sub
        End If
        For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
            CrystalReport2.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
        Next intIndex
        
'        For intIndex = UBound(Parameter) - LBound(Parameter) + 1 To 30
'             CrystalReport1.ParameterFields(intIndex) = ""
'        Next intIndex
        
        CrystalReport2.WindowShowGroupTree = True
        CrystalReport2.WindowShowSearchBtn = True
        CrystalReport2.WindowState = crptMaximized
        ODBCSetting clsArya.ServerName, clsArya.DbName
        CrystalReport2.Connect = CrystallConnection
        CrystalReport2.Action = 1
        CrystalReport2.RetrieveDataFiles
        
        If Screen.Width > 12000 Then
            CrystalReport2.PageZoom (100)
        Else
            CrystalReport2.PageZoom (75)
        End If
    
    End If
    
    

Exit Sub

Err_Handler:
    LogSaveNew "frmBenefit => ", err.Description, err.Number, err.Source, "Printing"
    ShowErrorMessage
    err.Clear

End Sub

Private Sub CalculateTotalLabels()
    On Error GoTo Err_Handler
    
    If cmbInventory.ListIndex = -1 Or cmbSalMali.ListIndex = -1 Then Exit Sub
    
    MyFormAddEditMode = EnumAddEditMode.ViewMode 'VIEW Mode
    
    SetFirstToolBar
    
    vsGood.Rows = 1
    
    If lstGoodLevel1.ListCount < 1 Then Exit Sub
    If lstGoodLevel2.ListCount < 1 Then Exit Sub
    
    Dim i As Long
    Dim j As Long
    Dim intSelectedLevel1 As Integer
    Dim intSelectedLevel2 As Integer
    Dim strSelectedLevels As String
    
    intSelectedLevel1 = -1
    intSelectedLevel2 = -1
    
    For i = 0 To lstGoodLevel1.ListCount - 1
        If lstGoodLevel1.Selected(i) = True Then
            intSelectedLevel1 = i
        End If
    Next i
    
    strSelectedLevels = ""
    For i = 0 To lstGoodLevel2.ListCount - 1
        If lstGoodLevel2.Selected(i) = True Then
            intSelectedLevel2 = i
            strSelectedLevels = strSelectedLevels + "," + CStr(lstGoodLevel2.ItemData(i))
        End If
    Next i
        
    Dim level1 As Integer
    If intSelectedLevel1 <> -1 And intSelectedLevel2 <> -1 Then
        level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
        strSelectedLevels = Right(strSelectedLevels, Len(strSelectedLevels) - 1)
    ElseIf intSelectedLevel1 <> -1 And intSelectedLevel2 = -1 Then  'Or intSelectedLevel2 = -1
       level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
       strSelectedLevels = ""
    Else
        strSelectedLevels = ""
        level1 = -1
    End If
    
    
    TotalBenefitLoss = 0
    TotalFirstPrice = 0
    TotalBuyAmount = 0
    TotalBuyReturnAmount = 0
    TotalMojodiPrice = 0
    TotalSellReturnAmount = 0
    TotalSellAmount = 0
    TotalHavalehAmount = 0
    TotalResidAmount = 0
    TotalLossAmount = 0
    TotalFinalSellAmount = 0
    TotalFinalSellReturnAmount = 0
    
    SaleDiscountTotal = 0
    
    LblTotalFirstPrice.Caption = Format(TotalFirstPrice, "#,##")
    LblTotalBuyAmount.Caption = Format(TotalBuyAmount, "#,##")
    LblTotalBuyReturnAmount.Caption = Format(TotalBuyReturnAmount, "#,##")
    LblTotalLossAmount.Caption = Format(TotalLossAmount, "#,##")
    LblTotalhavalehAmount.Caption = Format(TotalHavalehAmount, "#,##")
    LblTotalResidAmount.Caption = Format(TotalResidAmount, "#,##")
    LblTotalMojodiPrice.Caption = Format(TotalMojodiPrice, "#,##")
    LblTotalSellAmount.Caption = Format(TotalSellAmount, "#,##")
    LblTotalSellReturnAmount.Caption = Format(TotalSellReturnAmount, "#,##")
    lblTotalFinalSellAmount.Caption = Format(TotalFinalSellAmount, "#,##")
    LblTotalFinalSellReturnAmount.Caption = Format(TotalFinalSellReturnAmount, "#,##")
    LblTotalBenefitLoss.Caption = Format(TotalBenefitLoss, "#,##")
    
    Me.MousePointer = vbHourglass
    DoEvents
    Dim L_Rst As New ADODB.Recordset
    
    ReDim Parameter(5)
    Parameter(0) = GenerateInputParameter("@DateBefore", adVarWChar, 8, Me.txtDateFrom.Text)
    Parameter(1) = GenerateInputParameter("@DateAfter", adVarWChar, 8, Me.txtDateTo.Text)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(3) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(4) = GenerateInputParameter("@GoodLevel1", adInteger, 4, level1)
    Parameter(5) = GenerateInputParameter("@SelectedLevelsString", adVarWChar, 4000, strSelectedLevels)
    
    Set L_Rst = RunParametricStoredProcedure2Rec("Get_Benefit_Loss", Parameter)
    
    If L_Rst.BOF = True And L_Rst.EOF = True Then
        LblTotalFirstPrice.Caption = ""
        LblTotalBuyAmount.Caption = ""
        LblTotalBuyReturnAmount.Caption = ""
        LblTotalhavalehAmount.Caption = ""
        LblTotalResidAmount.Caption = ""
        LblTotalLossAmount.Caption = ""
        LblTotalMojodiPrice.Caption = ""
        LblTotalSellAmount.Caption = ""
        lblTotalFinalSellAmount.Caption = ""
        LblTotalFinalSellReturnAmount.Caption = ""
        LblTotalBenefitLoss.Caption = ""
        SaleDiscountTotal = 0
        Me.MousePointer = vbDefault
        Exit Sub
        Set L_Rst = Nothing
    End If
    
    i = 1
    With vsGood
        FWProgressBar1.Value = 0
        
        Static arr As Variant
        arr = L_Rst.GetRows
        .LoadArray arr
                
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
            TotalFirstPrice = TotalFirstPrice + Val(.TextMatrix(i, 6))
            TotalBuyAmount = TotalBuyAmount + Val(.TextMatrix(i, 7))
            TotalSellAmount = TotalSellAmount + Val(.TextMatrix(i, 15))
            TotalSellReturnAmount = TotalSellReturnAmount + Val(.TextMatrix(i, 17))
            TotalBuyReturnAmount = TotalBuyReturnAmount + Val(.TextMatrix(i, 8))
            .TextMatrix(i, 14) = Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 7)) - Val(.TextMatrix(i, 8)) - Val(.TextMatrix(i, 10)) + Val(.TextMatrix(i, 11))
            If Val(.TextMatrix(i, 12)) > 0 Then .TextMatrix(i, 13) = Val(.TextMatrix(i, 14)) / Val(.TextMatrix(i, 12))
            TotalMojodiPrice = TotalMojodiPrice + Val(.TextMatrix(i, 14))
            TotalHavalehAmount = TotalHavalehAmount + Val(.TextMatrix(i, 10))
            TotalResidAmount = TotalResidAmount + Val(.TextMatrix(i, 11))
            TotalLossAmount = TotalLossAmount + Val(.TextMatrix(i, 9))
            TotalFinalSellAmount = TotalFinalSellAmount + Val(.TextMatrix(i, 16))
            TotalFinalSellReturnAmount = TotalSellReturnAmount + Val(.TextMatrix(i, 18))
            SaleDiscountTotal = IIf(IsNull(.TextMatrix(i, 20)), 0, Val(.TextMatrix(i, 20)))
            
            .TextMatrix(i, 4) = Val(Format(.TextMatrix(i, 4), "##.000"))
            .TextMatrix(i, 12) = Val(Format(.TextMatrix(i, 12), "##.000"))
        Next
        .Cols = 20
        
        Me.MousePointer = vbDefault
        FWProgressBar1.Value = 0
        L_Rst.Close
        Set L_Rst = Nothing
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        If .Rows > 1 Then
            .Cell(flexcpAlignment, 1, 2, .Rows - 1, 2) = flexAlignRightCenter
        End If
        
'        .AutoSizeMode = flexAutoSizeColWidth
'        .AutoSize 1, 14
    End With
    TotalSellAmount = TotalSellAmount - SaleDiscountTotal
    TotalBenefitLoss = (TotalSellAmount - TotalSellReturnAmount) - (TotalFinalSellAmount - TotalFinalSellReturnAmount)
    
    LblTotalFirstPrice.Caption = Format(TotalFirstPrice, "#,##")
    LblTotalBuyAmount.Caption = Format(TotalBuyAmount, "#,##")
    LblTotalBuyReturnAmount.Caption = Format(TotalBuyReturnAmount, "#,##")
    LblTotalLossAmount.Caption = Format(TotalLossAmount, "#,##")
    LblTotalhavalehAmount.Caption = Format(TotalHavalehAmount, "#,##")
    LblTotalResidAmount.Caption = Format(TotalResidAmount, "#,##")
    LblTotalMojodiPrice.Caption = Format(TotalMojodiPrice, "#,##") & clsArya.UnitPrice
    LblTotalSellAmount.Caption = Format(TotalSellAmount, "#,##")
    lblTotalFinalSellAmount.Caption = Format(TotalFinalSellAmount, "#,##")
    lblTotalFinalSellAmount.Caption = "(" & lblTotalFinalSellAmount.Caption & ")"
    LblTotalSellReturnAmount.Caption = Format(TotalSellReturnAmount, "#,##")
    LblTotalSellReturnAmount.Caption = "(" & LblTotalSellReturnAmount.Caption & ")"
    LblTotalFinalSellReturnAmount.Caption = Format(TotalFinalSellReturnAmount, "#,##")
    LblTotalBenefitLoss.Caption = Format(TotalBenefitLoss, "#,##") & clsArya.UnitPrice

    FillSanadNo
    
Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmBenefit => ", err.Description, err.Number, err.Source, "CalculateTotalLabels"
    Me.MousePointer = vbDefault
    
End Sub
Private Sub FillSanadNo()
    Set rs = RunStoredProcedure2RecordSet("Get_All_tAccountYears")
    Do While rs.EOF = False
        If rs!AccountYear = cmbSalMali.Text Then
            txtSanadNo = IIf(IsNull(rs!FirstMojodi), "", rs!FirstMojodi)
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close

End Sub
Private Sub SetTooltipText()
    
    With FWToolTip
        .BackColor = vbYellow
        .Ballon = True
        .Margin(flwToolTipMarginLeft) = 20
        .MaxWidth = 500
        .DelayTime(flwToolTipDelayDefault) = 100
        '.DelayTime(flwToolTipDelayInitial) = 100
        .DelayTime(flwToolTipDelayShow) = 3000
        .DelayTime(flwToolTipDelayReshow) = 1500
'        .Text(frmBenefit) = "›—„Ê· „Õ«”»Â ”Êœ- “Ì«‰ ﬂ·:" & vbCrLf & vbCrLf & "((ﬂ· ›—Ê‘ „‰Â«Ì ﬂ· »—ê‘  «“ ›—Ê‘)-(ﬂ· »Â«¡  „«„ ‘œÂ ›—Ê‘ Ê » «“ ›—Ê‘))"
        .Text(vsGood) = "»« ò·Ìò —ÊÌ ‰«„ ò«·« ê—œ‘ ò«·« œ— «‰»«— ‰„«Ì‘ œ«œ«Â „Ì ‘Êœ" & vbCrLf & vbCrLf & "œ— ’Ê— Ì ﬂÂ «‰œ«“Â Â«Ì ” Ê‰ Â« —«  €ÌÌ— œÂÌœ°  €ÌÌ—«  »—«Ì ‰„«Ì‘ »⁄œÌ À»  ŒÊ«Âœ ‘œ"
        .Text(CmdUpdateHavaleh_Resid) = "»Â —Ê“ —”«‰Ì ﬁÌ„  ÕÊ«·Â Ê —”Ìœ »—«Ì «‰»«— " & cmbInventory.Text & " Å” «“ „Õ«”»Â „»·€ „Ì«‰êÌ‰ „Ê“Ê‰"
        .Text(CmdUpdateHavaleh_Resid) = .Text(CmdUpdateHavaleh_Resid) & vbCrLf & vbCrLf & "„»·€ „Ì«‰êÌ‰ „Ê“Ê‰ »—«»— «”  »« ‰”»  „»·€ ﬂ· Œ—Ìœ »Â  ⁄œ«œ ﬂ·"
        .Text(StoreDataUpdate) = "„Õ«”»Â ﬁÌ„   „«„ ‘œÂ ﬂ«·« œ— «‰»«— " & cmbInventory.Text & " »« «” ›«œÂ «“ —Ê‘ „Ì«‰êÌ‰ „Ê“Ê‰"
        .Text(StoreDataUpdate) = .Text(StoreDataUpdate) & vbCrLf & vbCrLf & "„»·€ „Ì«‰êÌ‰ „Ê“Ê‰ »—«»— «”  »« ‰”»  „»·€ ﬂ· Œ—Ìœ »Â  ⁄œ«œ ﬂ· Œ—Ìœ"
'        .Text(ChkZeroNegative) = "œ— ’Ê— Ì ﬂÂ »Â ⁄·  Ê«—œ ‰ﬂ—œ‰ Œ—Ìœ Ê „ÊÃÊœÌ «Ê·ÌÂ ﬁÌ„   „«„ ‘œÂ „‰›Ì „Õ«”»Â ê—œœ° ¬‰ —« ’›— œ— ‰Ÿ— „Ì êÌ—œ"
       ' .Text(LblTotalMojodiPrice) = "Õ«’·÷—» „ÊÃÊœÌ œ— ﬁÌ„   „«„ ‘œÂ"
        .Text(LblTotalMojodiPrice) = "Ã„⁄ Ê—ÊœÌ Â« „‰Â«Ì Œ—ÊÃÌ Â«"
       End With
End Sub

Public Function FindGood(ByVal GoodCode As Long) As Boolean
    Dim index As Long
    index = -1
    With vsGood
        index = .FindRow(GoodCode, -1, IdxColGoodCode, False, True)
        If index > 0 Then
            .ShowCell index, IdxColGoodCode
            FindGood = True
        Else
            FindGood = False
        End If
    End With
End Function



