VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmStationsetting 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14685
   Icon            =   "frmStationsetting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10170
   ScaleWidth      =   14685
   Begin VB.ComboBox cboStations 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3540
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   261
      Top             =   60
      Width           =   2115
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmStationsetting.frx":A4C2
      TabIndex        =   1
      Top             =   0
      Width           =   480
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9525
      Left            =   255
      TabIndex        =   2
      Top             =   585
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   16801
      _Version        =   393216
      MousePointer    =   3
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ÅÌ‘ ›—÷ Â«Ì ‰„«Ì‘ Ê Ã” ÃÊ"
      TabPicture(0)   =   "frmStationsetting.frx":A548
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame53"
      Tab(0).Control(1)=   "Picture1"
      Tab(0).Control(2)=   "Picture2"
      Tab(0).Control(3)=   "Picture3"
      Tab(0).Control(4)=   "Picture4"
      Tab(0).Control(5)=   "Frame25"
      Tab(0).Control(6)=   "Frame14"
      Tab(0).Control(7)=   "Picture5"
      Tab(0).Control(8)=   "Picture6"
      Tab(0).Control(9)=   "Picture7"
      Tab(0).Control(10)=   "Picture8"
      Tab(0).Control(11)=   "Picture9"
      Tab(0).Control(12)=   "Picture10"
      Tab(0).Control(13)=   "Picture11"
      Tab(0).Control(14)=   "Picture12"
      Tab(0).Control(15)=   "Picture13"
      Tab(0).Control(16)=   "Picture14"
      Tab(0).Control(17)=   "Picture16"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Ê”«Ì· Ã«‰»Ì"
      TabPicture(1)   =   "frmStationsetting.frx":A564
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame12"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame23"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame_CallerId"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame4"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame7"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame6"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame9"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame10"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "”«Ì— „Ê«—œ"
      TabPicture(2)   =   "frmStationsetting.frx":A580
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame_Control"
      Tab(2).Control(2)=   "Frame_Rate"
      Tab(2).Control(3)=   "Frame20"
      Tab(2).Control(4)=   "Frame44"
      Tab(2).Control(5)=   "Frame43"
      Tab(2).Control(6)=   "Frame19"
      Tab(2).Control(7)=   "Frame16"
      Tab(2).Control(8)=   "Frame_Kala"
      Tab(2).ControlCount=   9
      Begin VB.Frame Frame10 
         Caption         =   "”«“„«‰Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3900
         Left            =   375
         RightToLeft     =   -1  'True
         TabIndex        =   257
         Top             =   3990
         Width           =   4785
         Begin VB.TextBox txtDevice2Id 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   272
            Text            =   "2"
            Top             =   3360
            Width           =   645
         End
         Begin VB.TextBox txtDevice2IP 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   269
            Text            =   "192.168.1.101"
            Top             =   2910
            Width           =   2130
         End
         Begin VB.TextBox txtListFont 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   375
            RightToLeft     =   -1  'True
            TabIndex        =   266
            Text            =   "14"
            Top             =   2490
            Width           =   645
         End
         Begin VB.TextBox txtDeviceId 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   390
            RightToLeft     =   -1  'True
            TabIndex        =   265
            Text            =   "1"
            Top             =   2055
            Width           =   645
         End
         Begin VB.TextBox txtDeviceIP 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   390
            RightToLeft     =   -1  'True
            TabIndex        =   263
            Text            =   "192.168.1.100"
            Top             =   1155
            Width           =   2115
         End
         Begin VB.CheckBox chkListAutoLoad 
            Alignment       =   1  'Right Justify
            Caption         =   "»«—ê–«—Ì « Ê„« Ìò ·Ì”  «‰ Ÿ«—"
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
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   262
            Top             =   825
            Width           =   4095
         End
         Begin VB.TextBox txtPersonIdRefreshTime 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   405
            RightToLeft     =   -1  'True
            TabIndex        =   259
            Text            =   "5"
            Top             =   1650
            Width           =   645
         End
         Begin VB.CheckBox chkPersonIdCheck 
            Alignment       =   1  'Right Justify
            Caption         =   "«” ›«œÂ «“ œ” ê«Â  ‘ŒÌ’ ÂÊÌ  PWxxx  "
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
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   258
            Top             =   450
            Width           =   4095
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "¬Ì œÌ œ” ê«Â œÊ„"
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
            Left            =   1755
            TabIndex        =   271
            Top             =   3315
            Width           =   2550
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "¬Ì ÅÌ œ” ê«Â œÊ„"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2580
            TabIndex        =   270
            Top             =   2985
            Width           =   1710
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "”«Ì“ ›Ê‰  ·Ì” "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1485
            TabIndex        =   268
            Top             =   2535
            Width           =   2550
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "¬Ì œÌ œ” ê«Â"
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
            Left            =   1485
            TabIndex        =   267
            Top             =   2130
            Width           =   2550
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "¬Ì ÅÌ œ” ê«Â"
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
            Left            =   3045
            TabIndex        =   264
            Top             =   1230
            Width           =   1215
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "“„«‰ ŒÊ«‰œ‰ «“ œ” ê«Â - À«‰ÌÂ"
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
            Left            =   1515
            TabIndex        =   260
            Top             =   1725
            Width           =   2550
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "«—”«· ÅÌ«„ò"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1065
         Left            =   405
         RightToLeft     =   -1  'True
         TabIndex        =   254
         Top             =   450
         Width           =   4740
         Begin VB.CheckBox chkAryaSmsPanel 
            Alignment       =   1  'Right Justify
            Caption         =   "Ê—Êœ « Ê„« Ìò »Â «—”«· ÅÌ«„ò"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   660
            RightToLeft     =   -1  'True
            TabIndex        =   255
            Top             =   420
            Width           =   3300
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "ﬂ«—  ŒÊ«‰ òÌ»Ê—œÌ"
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
         Height          =   930
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   240
         Top             =   8175
         Width           =   4395
         Begin VB.TextBox txtStartCharacter 
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
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   241
            Top             =   405
            Width           =   645
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "ò«—«ò — ‘—Ê⁄ "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1560
            TabIndex        =   242
            Top             =   435
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "ﬂ«—  ŒÊ«‰ „«Ì›—"
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
         Height          =   2475
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   235
         Top             =   1515
         Width           =   4815
         Begin VB.CheckBox chkLoyaltyAllCustomers 
            Alignment       =   1  'Right Justify
            Caption         =   " ”Ì” „ Ê›«œ«—Ì »—«Ì „‘ —Ì«‰ „ ›—ﬁÂ"
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
            Left            =   615
            RightToLeft     =   -1  'True
            TabIndex        =   253
            Top             =   2040
            Width           =   3675
         End
         Begin VB.CheckBox chkLoyaltyCustomers 
            Alignment       =   1  'Right Justify
            Caption         =   "”Ì” „ Ê›«œ«—Ì „‘ —Ì«‰ œ«—«Ì ò«— "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   252
            Top             =   1635
            Width           =   4215
         End
         Begin VB.TextBox txtRfidInterval 
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
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   238
            Text            =   "1000"
            Top             =   1200
            Width           =   645
         End
         Begin VB.CheckBox chkRfidReader 
            Alignment       =   1  'Right Justify
            Caption         =   "ò«— ŒÊ«‰ ›⁄«·"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   237
            Top             =   480
            Width           =   2415
         End
         Begin VB.CheckBox ChkRfidLongBeep 
            Alignment       =   1  'Right Justify
            Caption         =   "»Êﬁ «Œÿ«— œ— “„«‰ ⁄œ„ ‘‰«”«ÌÌ ò«— "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   236
            Top             =   840
            Width           =   3975
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            Caption         =   "“„«‰ ŒÊ«‰œ‰ »Â „Ì·Ì À«‰ÌÂ"
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
            Left            =   1560
            TabIndex        =   239
            Top             =   1200
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   -69480
         TabIndex        =   227
         Top             =   4800
         Width           =   3855
         Begin VB.CheckBox chkShowOption 
            Alignment       =   1  'Right Justify
            Caption         =   "‰„«Ì‘ ¬Å‘‰ ò«·«Â« —ÊÌ „‰Ê Â«Ì  òÌ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   229
            Top             =   600
            Width           =   3495
         End
         Begin VB.CheckBox chkHasOptionPrice 
            Alignment       =   1  'Right Justify
            Caption         =   "„Õ«”»Â ﬁÌ„  ¬Å‘‰ Â« œ— ›«ò Ê—"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   228
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "ŒÊœ Å—œ«“ »«‰ﬂÌ"
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
         Height          =   1440
         Left            =   10050
         RightToLeft     =   -1  'True
         TabIndex        =   211
         Top             =   7635
         Width           =   4035
         Begin VB.CheckBox ChkPosPayment 
            Alignment       =   1  'Right Justify
            Caption         =   "Å—œ«Œ  « Ê„« Ìﬂ »Ê”Ì·Â ŒÊœÅ—œ«“"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   330
            RightToLeft     =   -1  'True
            TabIndex        =   213
            Top             =   435
            Width           =   3345
         End
         Begin VB.ComboBox CmbPosModel 
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
            Left            =   90
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   212
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label73 
            Alignment       =   1  'Right Justify
            Caption         =   "‰Ê⁄ ŒÊœÅ—œ«“ »«‰òÌ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2250
            TabIndex        =   214
            Top             =   870
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Å‘ Ì»«‰ êÌ—Ì —ÊÌ œÌ «»Ì”"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1005
         Left            =   5385
         RightToLeft     =   -1  'True
         TabIndex        =   209
         Top             =   7170
         Width           =   4395
         Begin VB.CheckBox ChkAutoBackUp 
            Alignment       =   1  'Right Justify
            Caption         =   "Å‘ Ì»«‰ Â‰ê«„ Œ—ÊÃ «“ ”Ì” „"
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
            Left            =   765
            RightToLeft     =   -1  'True
            TabIndex        =   210
            Top             =   495
            Width           =   3435
         End
      End
      Begin VB.PictureBox Picture16 
         Height          =   735
         Left            =   -74640
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   195
         Top             =   3360
         Width           =   6615
         Begin VB.OptionButton OptAlphabetGoodSearch 
            Alignment       =   1  'Right Justify
            Caption         =   "« Œ«» ﬂ«·« »« Õ—Ê› «·›»« Ê ‰«„ ﬂ«·«"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   197
            Top             =   120
            Width           =   2895
         End
         Begin VB.OptionButton OptAlphabetGoodSearch 
            Alignment       =   1  'Right Justify
            Caption         =   "« ‰Œ«» ﬂ«·« »« Õ—Ê› «·›»« Ê ‘„«—Â —œÌ›"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3120
            TabIndex        =   196
            Top             =   120
            Width           =   3375
         End
      End
      Begin VB.PictureBox Picture14 
         Height          =   735
         Left            =   -74640
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   192
         Top             =   2640
         Width           =   6615
         Begin VB.OptionButton OptViewTempAddress 
            Alignment       =   1  'Right Justify
            Caption         =   "‰„«Ì‘  Ê÷ÌÕ«  „‘ —ﬂ"
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
            Index           =   0
            Left            =   3750
            TabIndex        =   194
            Top             =   120
            Width           =   2535
         End
         Begin VB.OptionButton OptViewTempAddress 
            Alignment       =   1  'Right Justify
            Caption         =   "‰„«Ì‘ ¬œ—” „Êﬁ "
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
            Index           =   1
            Left            =   840
            TabIndex        =   193
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ﬂ«—  ŒÊ«‰ „€‰«ÿÌ”Ì Ê „«Ì›— „ ’· »Â ÅÊ— "
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
         Height          =   1380
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   185
         Top             =   7905
         Width           =   4815
         Begin VB.TextBox TxtNumberOfCardReader 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Text            =   "5"
            Top             =   855
            Width           =   645
         End
         Begin VB.TextBox TxtStartNumberCartReader 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   186
            Text            =   "1"
            Top             =   375
            Width           =   645
         End
         Begin VB.Label Label97 
            Alignment       =   1  'Right Justify
            Caption         =   " ⁄œ«œ «—ﬁ«„ »—«Ì ŒÊ«‰œ‰"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1560
            TabIndex        =   189
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label96 
            Alignment       =   1  'Right Justify
            Caption         =   "‘—Ê⁄ ŒÊ«‰œ‰ «“ —ﬁ„"
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
            Left            =   1920
            TabIndex        =   187
            Top             =   390
            Width           =   1695
         End
      End
      Begin VB.Frame Frame_Control 
         Caption         =   "ò‰ —·"
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
         Height          =   3705
         Left            =   -65400
         RightToLeft     =   -1  'True
         TabIndex        =   170
         Top             =   2880
         Width           =   4455
         Begin VB.CheckBox chkForceTax 
            Alignment       =   1  'Right Justify
            Caption         =   "«Ã»«—Ì ò—œ‰ ’œÊ— ›«ò Ê— —”„Ì"
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
            Left            =   450
            RightToLeft     =   -1  'True
            TabIndex        =   256
            Top             =   3210
            Width           =   3705
         End
         Begin VB.CheckBox chkOtherPartition 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄œ„ ‰„«Ì‘ ÊÀ»  „Ì“Â«Ì Å«— Ì‘‰ œÌê—"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   255
            RightToLeft     =   -1  'True
            TabIndex        =   222
            Top             =   2880
            Width           =   3900
         End
         Begin VB.CheckBox ChkAutoCashClose 
            Alignment       =   1  'Right Justify
            Caption         =   "»” ‰ ’‰œÊﬁ  —Ê“Â«Ì ﬁ»· »’Ê—  « Ê„« Ìﬂ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   206
            Top             =   2520
            Width           =   3990
         End
         Begin VB.TextBox TxtCountCustomerShiftBuy 
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
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   201
            Top             =   1200
            Width           =   1125
         End
         Begin VB.TextBox TxtCountCustomerDailyBuy 
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
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   174
            Top             =   720
            Width           =   1125
         End
         Begin VB.TextBox txtCountCustomerGoods 
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
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   173
            Top             =   1680
            Width           =   1125
         End
         Begin VB.CheckBox CheckTable 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   3720
            TabIndex        =   172
            Top             =   360
            Width           =   255
         End
         Begin VB.CheckBox ChkStopOnEditFich 
            Alignment       =   1  'Right Justify
            Caption         =   " Êﬁ› —ÊÌ ›«ò Ê— Â‰ê«„ ç«Å Ê «’·«Õ Ê  ”ÊÌÂ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   171
            Top             =   2160
            Width           =   4020
         End
         Begin VB.Label Label103 
            Alignment       =   1  'Right Justify
            Caption         =   " œ›⁄«  Œ—Ìœ „‘ —ò œ— ‘Ì› "
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
            Left            =   1320
            TabIndex        =   200
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            Caption         =   " ⁄œ«œ ò«·« œ— ›«ò Ê— „‘ —ò"
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
            Left            =   1680
            TabIndex        =   177
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label Label67 
            Alignment       =   1  'Right Justify
            Caption         =   " œ›⁄«  Œ—Ìœ „‘ —ò œ— —Ê“"
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
            Left            =   1680
            TabIndex        =   176
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            Caption         =   "ﬂ‰ —· „Ì“Â«Ì Å—"
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
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   175
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame_Rate 
         Caption         =   "‰—Œ"
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
         Height          =   5415
         Left            =   -74040
         RightToLeft     =   -1  'True
         TabIndex        =   155
         Top             =   3960
         Width           =   3855
         Begin VB.CheckBox ChkFixRateChange 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   3360
            TabIndex        =   205
            Top             =   1155
            Width           =   255
         End
         Begin VB.CheckBox ChkShiftRate 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   3360
            TabIndex        =   198
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox ChkCustomerFeeDatabase 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   3360
            TabIndex        =   183
            Top             =   2355
            Width           =   255
         End
         Begin VB.CheckBox ChkMultiPrice 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   3360
            TabIndex        =   160
            Top             =   285
            Width           =   255
         End
         Begin VB.ComboBox CmbFromStoreFee 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   159
            Top             =   3120
            Width           =   1995
         End
         Begin VB.ComboBox CmbCustomerRate 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   158
            Top             =   2640
            Width           =   1995
         End
         Begin VB.CheckBox ChkUpdateBuyPrice 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   3360
            TabIndex        =   157
            Top             =   1485
            Width           =   255
         End
         Begin VB.CheckBox ChkUpdateSellPrice 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   3360
            TabIndex        =   156
            Top             =   1890
            Width           =   255
         End
         Begin FLWCtrls.FWNumericTextBox FWNumericPrice 
            Height          =   495
            Left            =   120
            TabIndex        =   161
            Top             =   3600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            Max             =   6
            Min             =   1
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FLWCtrls.FWNumericTextBox FWMaxNumericPrice 
            Height          =   495
            Left            =   120
            TabIndex        =   162
            Top             =   4800
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            Max             =   6
            Min             =   1
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FLWCtrls.FWNumericTextBox FWNumericOutPrice 
            Height          =   495
            Left            =   120
            TabIndex        =   190
            Top             =   4200
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            Max             =   6
            Min             =   1
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label105 
            Alignment       =   1  'Right Justify
            Caption         =   "À«»  „«‰œ‰ ‰—Œ »⁄œ «“  €ÌÌ—"
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
            Left            =   360
            TabIndex        =   204
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label102 
            Alignment       =   1  'Right Justify
            Caption         =   "«‰ Œ«» ‰—Œ »— «”«” ‘Ì› "
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
            Left            =   360
            TabIndex        =   199
            Top             =   660
            Width           =   2895
         End
         Begin VB.Label Label100 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷  ‰—Œ ﬂ«·«œ— Õ«·  »Ì—Ê‰"
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
            Left            =   1080
            TabIndex        =   191
            Top             =   4200
            Width           =   2655
         End
         Begin VB.Label Label95 
            Alignment       =   1  'Right Justify
            Caption         =   "«‰ Œ«» ›Ì „‘ —ò «“ œÌ «»Ì”"
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
            Left            =   120
            TabIndex        =   184
            Top             =   2280
            Width           =   3135
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "„«ﬂ“Ì„„ ‰—Œ ﬂ«·«"
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
            Left            =   1920
            TabIndex        =   169
            Top             =   4800
            Width           =   1815
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷  ‰—Œ ﬂ«·«œ— Õ«·  ⁄«œÌ"
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
            Left            =   1080
            TabIndex        =   168
            Top             =   3720
            Width           =   2655
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            Caption         =   "«‰ Œ«» ç‰œ ‰—ŒÌ"
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
            Left            =   1680
            TabIndex        =   167
            Top             =   225
            Width           =   1575
         End
         Begin VB.Label Label64 
            Alignment       =   1  'Right Justify
            Caption         =   " ⁄ÌÌ‰ ›Ì ÕÊ«·Â"
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
            Left            =   2280
            TabIndex        =   166
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            Caption         =   " ⁄ÌÌ‰ ›Ì „‘ —òÌ‰"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   165
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "»—Ê“ —”«‰Ì ›Ì Œ—Ìœ »⁄œ «“ ›Ì œ·ŒÊ«Â"
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
            Left            =   120
            TabIndex        =   164
            Top             =   1395
            Width           =   3135
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "»—Ê“ —”«‰Ì ›Ì ›—Ê‘ »⁄œ «“ ›Ì œ·ŒÊ«Â"
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
            Left            =   120
            TabIndex        =   163
            Top             =   1800
            Width           =   3255
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "‰„«Ì‘"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4455
         Left            =   -69480
         RightToLeft     =   -1  'True
         TabIndex        =   146
         Top             =   360
         Width           =   3855
         Begin VB.CheckBox ChkTouchScreen 
            Alignment       =   1  'Right Justify
            Caption         =   "‰„«Ì‘ ò·Ìœ ò«·«  «ç"
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
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   232
            Top             =   2625
            Width           =   2175
         End
         Begin VB.CheckBox chkOneClickShow 
            Alignment       =   1  'Right Justify
            Caption         =   "‰„«Ì‘ ›—„ Â« »«  ò ò·Ìò"
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
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   231
            Top             =   2280
            Width           =   2655
         End
         Begin VB.CheckBox ChkFastCustSave 
            Alignment       =   1  'Right Justify
            Caption         =   "›—„ À»  ”—Ì⁄ „‘ —òÌ‰"
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
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   219
            ToolTipText     =   "œ— ›—„ Ã” ÃÊÌ „‘ —Ì«‰ ›—„ À»  ”—Ì⁄ „‘Œ’«  „‘ —Ì«‰ »«“ „Ì ‘Êœ"
            Top             =   1875
            Width           =   2295
         End
         Begin VB.CheckBox chkCashPayment 
            Alignment       =   1  'Right Justify
            Caption         =   " ›—„  ”ÊÌÂ ‰ﬁœÌ "
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
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   1485
            Width           =   1695
         End
         Begin VB.CheckBox ChkPayFactorView 
            Alignment       =   1  'Right Justify
            Caption         =   "‰„«Ì‘ ﬂ·Ìœ œ—Ì«› "
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
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CheckBox ChDelView 
            Alignment       =   1  'Right Justify
            Caption         =   "‰„«Ì‘  ⁄œ«œ «—”«·Ì Â«"
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
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   150
            Top             =   630
            Width           =   2295
         End
         Begin VB.CheckBox ChkMenuViewAfterGood 
            Alignment       =   1  'Right Justify
            Caption         =   "‰„«Ì‘ „‰Ê »⁄œ«“«‰ Œ«» ﬂ«·«"
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
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   225
            Width           =   2535
         End
         Begin VB.TextBox TxtInvoiceRows 
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
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Text            =   "7"
            Top             =   3480
            Width           =   885
         End
         Begin VB.TextBox TxtPurchaseRows 
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
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Text            =   "8"
            Top             =   3960
            Width           =   885
         End
         Begin FLWCtrls.FWNumericTextBox FWNoRowMenu 
            Height          =   495
            Left            =   360
            TabIndex        =   234
            Top             =   2880
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            Max             =   2
            Min             =   1
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   " ⁄œ«œ »Œ‘ Â«Ì „‰Ê ò«·«"
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
            Left            =   1320
            TabIndex        =   233
            Top             =   3015
            Width           =   2295
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   " ⁄œ«œ ”ÿ—Â«Ì ›«ò Ê—›—Ê‘"
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
            Left            =   1320
            TabIndex        =   154
            Top             =   3480
            Width           =   2295
         End
         Begin VB.Label Label59 
            Alignment       =   1  'Right Justify
            Caption         =   " ⁄œ«œ ”ÿ—Â«Ì ›«ò Ê— Œ—Ìœ"
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
            Left            =   1320
            TabIndex        =   153
            Top             =   3960
            Width           =   2295
         End
      End
      Begin VB.Frame Frame44 
         Caption         =   " ”ÊÌÂ ›«ò Ê— œ— Õ«· Â«Ì  ·›‰Ì"
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
         Height          =   1545
         Left            =   -69480
         RightToLeft     =   -1  'True
         TabIndex        =   136
         Top             =   7800
         Width           =   3855
         Begin VB.CheckBox ChkByPhoneTableBalance 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   142
            Top             =   1080
            Width           =   255
         End
         Begin VB.CheckBox ChkByPhoneTablePayment 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   141
            Top             =   1080
            Width           =   255
         End
         Begin VB.CheckBox ChkByPhoneDeliveryBalance 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   140
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox ChkByPhoneDeliveryPayment 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2160
            TabIndex        =   139
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox ChkByPhoneSalonBalance 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   138
            Top             =   360
            Width           =   255
         End
         Begin VB.CheckBox ChkByPhoneSalonPayment 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2160
            TabIndex        =   137
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„Ì“"
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
            Left            =   2520
            TabIndex        =   145
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«—”«·Ì"
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
            Left            =   2520
            TabIndex        =   144
            Top             =   780
            Width           =   975
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "”«·‰"
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
            Left            =   2520
            TabIndex        =   143
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame Frame43 
         Caption         =   " ”ÊÌÂ ›«ò Ê— œ— Õ«· Â«Ì Õ÷Ê—Ì"
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
         Height          =   1920
         Left            =   -69480
         RightToLeft     =   -1  'True
         TabIndex        =   123
         Top             =   5880
         Width           =   3855
         Begin VB.CheckBox ChkInpersonTableBalance 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   131
            Top             =   1440
            Width           =   255
         End
         Begin VB.CheckBox ChkInpersonTablePayment 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   130
            Top             =   1440
            Width           =   255
         End
         Begin VB.CheckBox ChkInpersonOutBalance 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   129
            Top             =   1080
            Width           =   255
         End
         Begin VB.CheckBox ChkInpersonOutPayment 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   128
            Top             =   1080
            Width           =   255
         End
         Begin VB.CheckBox ChkInpersonDeliveryBalance 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   127
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox ChkInpersonDeliveryPayment 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   126
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox ChkInpersonSalonBalance 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   125
            Top             =   360
            Width           =   255
         End
         Begin VB.CheckBox ChkInpersonSalonPayment 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   124
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„Ì“"
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
            Left            =   2520
            TabIndex        =   135
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì—Ê‰"
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
            Left            =   2520
            TabIndex        =   134
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«—”«·Ì"
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
            Left            =   2400
            TabIndex        =   133
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "”«·‰"
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
            Left            =   2760
            TabIndex        =   132
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "ç«Å"
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
         Height          =   3615
         Left            =   -74040
         RightToLeft     =   -1  'True
         TabIndex        =   120
         Top             =   360
         Width           =   3975
         Begin VB.CheckBox chkLabelPrint 
            Alignment       =   1  'Right Justify
            Caption         =   "Å—Ì‰  ·Ì»· »« «” ›«œÂ «“ Å—›—«é"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   251
            Top             =   1680
            Width           =   3375
         End
         Begin VB.CheckBox chkPrintAfterOrder 
            Alignment       =   1  'Right Justify
            Caption         =   "ç«Å ›Ì‘ „‘ —Ì Ê ¬‘Å“Œ«‰Â »⁄œ «“ ”›«—‘"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   250
            Top             =   3240
            Width           =   3615
         End
         Begin VB.CheckBox ChkPrintAfterPayk 
            Alignment       =   1  'Right Justify
            Caption         =   "ç«Å ›«ﬂ Ê— ›—Ê‘ »⁄œ «“ «Œ ’«’ ÅÌò"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   249
            Top             =   2445
            Width           =   3375
         End
         Begin VB.CheckBox chkLableUsedGood 
            Alignment       =   1  'Right Justify
            Caption         =   "«” ›«œÂ «“ÃœÊ· „’—› »—«Ì  ⁄œ«œ ·Ì»· "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   230
            Top             =   2040
            Width           =   3495
         End
         Begin VB.CheckBox chkNotShowPrintNotice 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄œ„ ‰„«Ì‘ ÅÌ€«„ ›Ê—„  ç«Å  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   223
            Top             =   1280
            Width           =   3135
         End
         Begin VB.CheckBox ChkPrintAfterDeliver 
            Alignment       =   1  'Right Justify
            Caption         =   "ç«Å ›«ﬂ Ê— ›—Ê‘ »⁄œ «“  ÕÊÌ·"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   203
            Top             =   2835
            Width           =   2775
         End
         Begin VB.CheckBox ChChangeGoodPrint 
            Alignment       =   1  'Right Justify
            Caption         =   "ç«Å  €ÌÌ—«    ⁄œ«œ ﬂ«·«"
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
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   760
            Width           =   2295
         End
         Begin VB.CheckBox ChkPrintAfterTasvieh 
            Alignment       =   1  'Right Justify
            Caption         =   "ç«Å ›«ﬂ Ê— ›—Ê‘ »⁄œ «“  ”ÊÌÂ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame Frame16 
         Height          =   2775
         Left            =   -65400
         TabIndex        =   116
         Top             =   6600
         Width           =   4455
         Begin VB.CheckBox ChkEditCompatibleSamar1 
            Alignment       =   1  'Right Justify
            Caption         =   "«’·«Õ „‘«»Â ”„—1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   216
            Top             =   1464
            Width           =   2130
         End
         Begin VB.CheckBox ChkUndoRedoCompatibleSamar1 
            Alignment       =   1  'Right Justify
            Caption         =   "„—ÃÊ⁄ „‘«»Â ”„— 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1950
            RightToLeft     =   -1  'True
            TabIndex        =   215
            Top             =   1872
            Width           =   2220
         End
         Begin VB.CheckBox ChkUpDateCarryFee 
            Alignment       =   1  'Right Justify
            Caption         =   "»—Ê“ —”«‰Ì ﬂ—«ÌÂ Õ„· »⁄œ «“  €ÌÌ— œ— ›«ﬂ Ê—"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   345
            RightToLeft     =   -1  'True
            TabIndex        =   207
            Top             =   2280
            Width           =   3810
         End
         Begin VB.CheckBox ChkRefreshFichNo 
            Alignment       =   1  'Right Justify
            Caption         =   "»Â —Ê“ ﬂ—œ‰ ‘„«—Â ›Ì‘ œ— ‘»ﬂÂ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1065
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   648
            Width           =   3105
         End
         Begin VB.CheckBox ChkNoCurrentDay 
            Alignment       =   1  'Right Justify
            Caption         =   "’œÊ— ›Ì‘  «—ÌŒ €Ì— Ã«—Ì"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1605
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   1056
            Width           =   2565
         End
         Begin VB.CheckBox ChkCredit 
            Alignment       =   1  'Right Justify
            Caption         =   "„Õ«”»Â «⁄ »«— œ— ’Ê— Õ”«» „‘ —ﬂÌ‰"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   660
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   240
            Width           =   3510
         End
      End
      Begin VB.Frame Frame_Kala 
         Caption         =   "ò«·«"
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
         Height          =   2580
         Left            =   -65370
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   300
         Width           =   4335
         Begin VB.CheckBox ChkRepetitiveGood 
            Alignment       =   1  'Right Justify
            Caption         =   "ò«·«Ì  ò—«—Ì œ— ”ÿ— ÃœÌœ ›«ò Ê—Œ—Ìœ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   465
            TabIndex        =   221
            Top             =   2070
            Width           =   3405
         End
         Begin VB.CheckBox ChAlphabetic 
            Alignment       =   1  'Right Justify
            Caption         =   "«‰ Œ«» ﬂ«·« »« Õ—Ê› «·›»«"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1425
            TabIndex        =   111
            Top             =   1680
            Width           =   2445
         End
         Begin VB.CheckBox ChkNumberOfUnit 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3600
            TabIndex        =   110
            Top             =   1320
            Width           =   255
         End
         Begin VB.CheckBox ChkThreeSegmentSearch 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3600
            TabIndex        =   109
            Top             =   960
            Width           =   255
         End
         Begin VB.CheckBox ChkRowMojodiControl 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3600
            TabIndex        =   108
            Top             =   600
            Width           =   255
         End
         Begin VB.CheckBox ChkMojodiControlDefault 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   3600
            TabIndex        =   107
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "›—Ê‘   ⁄œ«œ œ— Ê«Õœ"
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
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   1410
            Width           =   2415
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "Ã” ÃÊÌ ”Â ”Ì·«»Ì"
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
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   1020
            Width           =   2415
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ ò‰ —· ò«·«œ—Â——œÌ›"
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
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   630
            Width           =   2775
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ ò‰ —· ò«·«"
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
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame_CallerId 
         Caption         =   "ò«·— ¬Ì œÌ"
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
         Height          =   6720
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   450
         Width           =   4455
         Begin VB.Frame Frame8 
            Caption         =   "ò«·—¬ÌœÌ »« Telnet"
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
            Height          =   1935
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   243
            Top             =   4635
            Width           =   4215
            Begin VB.CheckBox chkTelNetServerActive 
               Alignment       =   1  'Right Justify
               Caption         =   "«” ›«œÂ «“”—Ê—  Telnet »—«Ì ò«·—¬ÌœÌ"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   248
               Top             =   360
               Width           =   3735
            End
            Begin VB.TextBox txtTelnetServerPort 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   246
               Text            =   "2001"
               Top             =   1320
               Width           =   1005
            End
            Begin VB.TextBox txtTelnetServerIP 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   244
               Text            =   "192.168.1.1"
               Top             =   840
               Width           =   2565
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               Caption         =   "ÅÊ—   · ‰ "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2520
               TabIndex        =   247
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               Caption         =   "¬Ì ÅÌ ”—Ê—"
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
               TabIndex        =   245
               Top             =   840
               Width           =   1215
            End
         End
         Begin VB.CheckBox ChkfrmCallerId 
            Alignment       =   1  'Right Justify
            Caption         =   "»«“ ‘œ‰ « Ê„« Ìò ›—„ ò«·— ¬ÌœÌ"
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
            Left            =   750
            RightToLeft     =   -1  'True
            TabIndex        =   208
            Top             =   1920
            Width           =   3465
         End
         Begin VB.CheckBox ChkNetworkCallerId 
            Alignment       =   1  'Right Justify
            Caption         =   "«‘ —«ﬂ ŒÿÊÿ œ— ‘»òÂ"
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
            Left            =   825
            RightToLeft     =   -1  'True
            TabIndex        =   182
            Top             =   1440
            Width           =   3390
         End
         Begin VB.TextBox TxtResponsePort 
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
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   181
            Text            =   "4112"
            Top             =   4230
            Width           =   1005
         End
         Begin VB.TextBox TxtDiscoveryPort 
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
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   180
            Text            =   "4111"
            Top             =   3780
            Width           =   1005
         End
         Begin VB.CheckBox ChkAutoCallerId 
            Alignment       =   1  'Right Justify
            Caption         =   "»«“ ‘œ‰ « Ê„« Ìò  Ã” ÃÊÌ „‘ —ò "
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
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   480
            Width           =   3735
         End
         Begin VB.TextBox TxtCallerIdSpace 
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
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   2310
            Width           =   1005
         End
         Begin VB.TextBox TxtCityCode 
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
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Text            =   "21"
            Top             =   2850
            Width           =   1005
         End
         Begin VB.TextBox TxtNumberOfId 
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
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Text            =   "8"
            Top             =   3345
            Width           =   1005
         End
         Begin VB.CheckBox ChkAlmLogFile 
            Alignment       =   1  'Right Justify
            Caption         =   "À»  Êﬁ«Ì⁄ —ÊÌ ÅÊ—   Alm_p1"
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
            Left            =   645
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   975
            Width           =   3585
         End
         Begin VB.Label Label92 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÊ—  œ—Ì«›  «ÿ·«⁄« "
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
            Left            =   2280
            TabIndex        =   179
            Top             =   4260
            Width           =   1935
         End
         Begin VB.Label Label91 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÊ—  «—”«· «ÿ·«⁄« "
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
            Left            =   2280
            TabIndex        =   178
            Top             =   3810
            Width           =   1935
         End
         Begin VB.Label Label72 
            Alignment       =   1  'Right Justify
            Caption         =   " ›«’·Â «Ê·Ì‰ —ﬁ„ ﬂ«·— ¬Ì œÌ œ—"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   1800
            TabIndex        =   105
            Top             =   2355
            Width           =   2415
         End
         Begin VB.Label Label86 
            Alignment       =   1  'Right Justify
            Caption         =   "Alm_3"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   104
            Top             =   2370
            Width           =   615
         End
         Begin VB.Label Label87 
            Alignment       =   1  'Right Justify
            Caption         =   "»œÊ‰ ’›— ) òœ ‘Â—” «‰  )"
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
            Left            =   1320
            TabIndex        =   103
            Top             =   2865
            Width           =   2895
         End
         Begin VB.Label Label88 
            Alignment       =   1  'Right Justify
            Caption         =   " ⁄œ«œ «—ﬁ«„ Œÿ  ·›‰"
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
            Left            =   1320
            TabIndex        =   102
            Top             =   3360
            Width           =   2895
         End
         Begin VB.Label Label89 
            Alignment       =   1  'Right Justify
            Caption         =   "À»  Êﬁ«Ì⁄ —ÊÌ ÅÊ—  "
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
            Left            =   1815
            TabIndex        =   101
            Top             =   975
            Width           =   1935
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Ê”«Ì· Ã«‰»Ì"
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
         Height          =   2895
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   2925
         Width           =   3975
         Begin VB.CheckBox ChkCustomerFarsi 
            Alignment       =   1  'Right Justify
            Caption         =   "‰„«Ì‘ê— ›«—”Ì"
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
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   1770
            Width           =   3015
         End
         Begin VB.CheckBox ChkCustomerAscii 
            Alignment       =   1  'Right Justify
            Caption         =   "‰„«Ì‘ê— «”òÌ òœ"
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
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   1260
            Width           =   3015
         End
         Begin VB.CheckBox ChkSoundAlarm 
            Alignment       =   1  'Right Justify
            Caption         =   "¬·«—„ ’Ê Ì"
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
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   840
            Width           =   2055
         End
         Begin VB.CheckBox ChOpenDrawer 
            Alignment       =   1  'Right Justify
            Caption         =   "»«“ ‘œ‰ « Ê„« Ìﬂ ﬂ‘Ê"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   480
            Width           =   2415
         End
         Begin VB.CheckBox ChkCustomerOnlinePrice 
            Alignment       =   1  'Right Justify
            Caption         =   "‰„«Ì‘ „»·€ ò«·« œ— Œÿ «Ê· ‰„«Ì‘ê—"
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
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   2265
            Width           =   3495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " —«“Ê"
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
         Height          =   1590
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   5805
         Width           =   3975
         Begin VB.ComboBox cmbTypeBascule 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   87
            Top             =   360
            Width           =   2235
         End
         Begin VB.CheckBox ChkDirectBascule 
            Alignment       =   1  'Right Justify
            Caption         =   " —«“ÊÌ „” ﬁÌ„"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   825
            Width           =   2055
         End
         Begin VB.CheckBox ChkTrazooBarcode 
            Alignment       =   1  'Right Justify
            Caption         =   "«‰ Œ«» »«—ﬂœ  —«“Ê"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   "‰Ê⁄ ·Ì»·  —«“Ê"
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
            Left            =   2400
            TabIndex        =   88
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "»«—òœ"
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
         Height          =   2445
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   480
         Width           =   3975
         Begin VB.CheckBox ChkBarcodeChance 
            Alignment       =   1  'Right Justify
            Caption         =   "»«—ﬂœ Ã«Ì“Â"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox ChkAutoBarcode 
            Alignment       =   1  'Right Justify
            Caption         =   "»«—ﬂœ « Ê„« Ìﬂ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtBarcodePrice 
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
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Text            =   "10000"
            Top             =   1920
            Width           =   1005
         End
         Begin VB.CheckBox ChkBarcodeAutoEscape 
            Alignment       =   1  'Right Justify
            Caption         =   "Œ—ÊÃ « Ê„« Ìﬂ «“ »«—ﬂœ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "»«—ﬂœ Ã«Ì“Â"
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
            Left            =   2160
            TabIndex        =   83
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "—Ì«·"
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
            Left            =   840
            TabIndex        =   82
            Top             =   1920
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture13 
         Height          =   855
         Left            =   -67560
         ScaleHeight     =   795
         ScaleWidth      =   6555
         TabIndex        =   73
         Top             =   7080
         Width           =   6615
         Begin VB.OptionButton OptcustOrder 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ÷Ê—Ì"
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
            Index           =   0
            Left            =   1920
            TabIndex        =   75
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton OptcustOrder 
            Alignment       =   1  'Right Justify
            Caption         =   " ·›‰Ì"
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
            Index           =   1
            Left            =   600
            TabIndex        =   74
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ ”›«—‘ œ— «‰ Œ«» „‘ —ﬂ"
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
            Left            =   3360
            TabIndex        =   76
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.PictureBox Picture12 
         Height          =   735
         Left            =   -67560
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   69
         Top             =   6360
         Width           =   6615
         Begin VB.OptionButton OptSearchFichDefault 
            Alignment       =   1  'Right Justify
            Caption         =   "„⁄„Ê·Ì"
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
            Index           =   0
            Left            =   2760
            TabIndex        =   71
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton OptSearchFichDefault 
            Alignment       =   1  'Right Justify
            Caption         =   "”Â —ﬁ„ ¬Œ—"
            CausesValidation=   0   'False
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
            Index           =   1
            Left            =   960
            TabIndex        =   70
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ Ã” ÃÊÌ ›Ì‘"
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
            Left            =   3840
            TabIndex        =   72
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.PictureBox Picture11 
         Height          =   735
         Left            =   -67560
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   65
         Top             =   5640
         Width           =   6615
         Begin VB.OptionButton OptStartForm 
            Alignment       =   1  'Right Justify
            Caption         =   "ÂÌçòœ«„"
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
            Index           =   2
            Left            =   120
            TabIndex        =   220
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton OptStartForm 
            Alignment       =   1  'Right Justify
            Caption         =   "›«ﬂ Ê—›—Ê‘"
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
            Index           =   0
            Left            =   2880
            TabIndex        =   67
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton OptStartForm 
            Alignment       =   1  'Right Justify
            Caption         =   "›—„ Œ—Ìœ"
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
            Index           =   1
            Left            =   1440
            TabIndex        =   66
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷  ‰„«Ì‘ «Ê·ÌÂ"
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
            Left            =   3840
            TabIndex        =   68
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.PictureBox Picture10 
         Height          =   735
         Left            =   -67560
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   60
         Top             =   4920
         Width           =   6615
         Begin VB.OptionButton OptcustServePlace 
            Alignment       =   1  'Right Justify
            Caption         =   "«—”«·Ì"
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
            Index           =   0
            Left            =   2160
            TabIndex        =   63
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton OptcustServePlace 
            Alignment       =   1  'Right Justify
            Caption         =   "”«·‰"
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
            Index           =   1
            Left            =   1200
            TabIndex        =   62
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton OptcustServePlace 
            Alignment       =   1  'Right Justify
            Caption         =   "„Ì“"
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
            Index           =   2
            Left            =   360
            TabIndex        =   61
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ „Õ· ”—Ê œ— «‰ Œ«» „‘ —ﬂ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   64
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.PictureBox Picture9 
         Height          =   735
         Left            =   -67560
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   56
         Top             =   4200
         Width           =   6615
         Begin VB.OptionButton OptDiscount 
            Alignment       =   1  'Right Justify
            Caption         =   "—ÊÌ ﬂ· ›«ﬂ Ê—"
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
            Index           =   0
            Left            =   2160
            TabIndex        =   58
            Top             =   120
            Width           =   1575
         End
         Begin VB.OptionButton OptDiscount 
            Alignment       =   1  'Right Justify
            Caption         =   "—ÊÌ ﬂ«·«Â«"
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
            Index           =   1
            Left            =   720
            TabIndex        =   57
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ Õ«·   Œ›Ì›"
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
            Left            =   3720
            TabIndex        =   59
            Top             =   120
            Width           =   2655
         End
      End
      Begin VB.PictureBox Picture8 
         Height          =   735
         Left            =   -67560
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   52
         Top             =   3480
         Width           =   6615
         Begin VB.OptionButton OptCustSearch 
            Alignment       =   1  'Right Justify
            Caption         =   "„⁄„Ê·Ì"
            CausesValidation=   0   'False
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
            Index           =   1
            Left            =   1440
            TabIndex        =   54
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton OptCustSearch 
            Alignment       =   1  'Right Justify
            Caption         =   "”—Ì⁄"
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
            Index           =   0
            Left            =   3000
            TabIndex        =   53
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ Ã” ÃÊÌ „‘ —ﬂ"
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
            Left            =   3960
            TabIndex        =   55
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.PictureBox Picture7 
         Height          =   735
         Left            =   -67560
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   47
         Top             =   2760
         Width           =   6615
         Begin VB.OptionButton OptSearch 
            Alignment       =   1  'Right Justify
            Caption         =   " „Ì“ ê—ÊÂÌ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   202
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton OptSearch 
            Alignment       =   1  'Right Justify
            Caption         =   "‘„«—Â ›Ì‘"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3240
            TabIndex        =   50
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton OptSearch 
            Alignment       =   1  'Right Justify
            Caption         =   "„Ì“"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2520
            TabIndex        =   49
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton OptSearch 
            Alignment       =   1  'Right Justify
            Caption         =   "«—”«·Ì"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1560
            TabIndex        =   48
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ Õ«·  Ã” ÃÊ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4560
            TabIndex        =   51
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture6 
         Height          =   735
         Left            =   -67560
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   43
         Top             =   2040
         Width           =   6615
         Begin VB.OptionButton OptPrint 
            Alignment       =   1  'Right Justify
            Caption         =   "›«ﬂ Ê— ›—Ê‘"
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
            Index           =   1
            Left            =   480
            TabIndex        =   45
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton OptPrint 
            Alignment       =   1  'Right Justify
            Caption         =   "ç«Å „Ãœœ"
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
            Index           =   0
            Left            =   2040
            TabIndex        =   44
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ Õ«·  ç«Å ›Ì‘ Â«Ì ﬁ»·Ì"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   46
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.PictureBox Picture5 
         Height          =   735
         Left            =   -67560
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   39
         Top             =   1320
         Width           =   6615
         Begin VB.OptionButton OptTable 
            Alignment       =   1  'Right Justify
            Caption         =   "›—„ ê«—”Ê‰"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   840
            TabIndex        =   41
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton OptTable 
            Alignment       =   1  'Right Justify
            Caption         =   "›«ﬂ Ê— ›—Ê‘"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   2400
            TabIndex        =   40
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ Õ«·  »«—ﬂœ „Ì“"
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
            Left            =   3840
            TabIndex        =   42
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.Frame Frame14 
         Height          =   1335
         Left            =   -67560
         TabIndex        =   36
         Top             =   7920
         Width           =   6615
         Begin VB.CheckBox chkFrame_Printers 
            Alignment       =   1  'Right Justify
            Caption         =   "ò‰ —· Å—Ì‰ —Â«Ì ‘»òÂ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   615
            TabIndex        =   226
            Top             =   240
            Width           =   2175
         End
         Begin VB.CheckBox ChkFinalCheck 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ ‰„«Ì‘ ›Ì‰«· çò"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3795
            TabIndex        =   38
            Top             =   240
            Width           =   2580
         End
         Begin VB.CheckBox ChkAutoTip 
            Alignment       =   1  'Right Justify
            Caption         =   "«‰⁄«„ « Ê„« Ìò"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4665
            TabIndex        =   37
            Top             =   720
            Width           =   1710
         End
      End
      Begin VB.Frame Frame25 
         Height          =   1215
         Left            =   -74640
         TabIndex        =   33
         Top             =   8040
         Width           =   6615
         Begin VB.CheckBox ChkTextIconViewH 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ ‰„«Ì‘ „ ‰Ì ¬ÌòÊ‰ Â«Ì ‰Ê«— «›ﬁÌ"
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
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   4815
         End
         Begin VB.CheckBox ChkInvoiceStatusDefault 
            Alignment       =   1  'Right Justify
            Caption         =   "‰„«Ì‘ ›«ò Ê— ›—Ê‘ »⁄œ «“ À»  ”›«—‘"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   600
            Width           =   4455
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   735
         Left            =   -67560
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   29
         Top             =   600
         Width           =   6615
         Begin VB.OptionButton OptDelivery 
            Alignment       =   1  'Right Justify
            Caption         =   "›—„ ÅÌﬂ "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   840
            TabIndex        =   31
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton OptDelivery 
            Alignment       =   1  'Right Justify
            Caption         =   "›«ﬂ Ê— ›—Ê‘"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   2280
            TabIndex        =   30
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ Õ«·  »«—ﬂœ «—”«·Ì"
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
            Left            =   3720
            TabIndex        =   32
            Top             =   120
            Width           =   2655
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   735
         Left            =   -74640
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   24
         Top             =   1920
         Width           =   6615
         Begin VB.OptionButton OptOrder 
            Alignment       =   1  'Right Justify
            Caption         =   "ò«„·"
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
            Index           =   2
            Left            =   240
            TabIndex        =   27
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton OptOrder 
            Alignment       =   1  'Right Justify
            Caption         =   "«‰Ã«„ ‘œÂ"
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
            Index           =   1
            Left            =   1200
            TabIndex        =   26
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton OptOrder 
            Alignment       =   1  'Right Justify
            Caption         =   "«‰Ã«„ ‰‘œÂ"
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
            Index           =   0
            Left            =   2640
            TabIndex        =   25
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ Ã” ÃÊÌ ”›«—‘"
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
            Left            =   3960
            TabIndex        =   28
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   735
         Left            =   -74640
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   20
         Top             =   1200
         Width           =   6615
         Begin VB.OptionButton OptDeletedGood 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ–› ‘Êœ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   1920
            TabIndex        =   22
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton OptDeletedGood 
            Alignment       =   1  'Right Justify
            Caption         =   "«’·«Õ ‘Êœ"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   480
            TabIndex        =   21
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            Caption         =   " ﬂ«·«Ì Ê“‰Ì Â‰ê«„ Õ–› œ— ›«ﬂ Ê—"
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
            Left            =   3600
            TabIndex        =   23
            Top             =   45
            Width           =   2775
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Left            =   -74640
         ScaleHeight     =   675
         ScaleWidth      =   6555
         TabIndex        =   16
         Top             =   480
         Width           =   6615
         Begin VB.OptionButton OptGoodSearch 
            Alignment       =   1  'Right Justify
            Caption         =   "”—Ì⁄"
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
            Index           =   0
            Left            =   2040
            TabIndex        =   18
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton OptGoodSearch 
            Alignment       =   1  'Right Justify
            Caption         =   "„⁄„Ê·Ì"
            CausesValidation=   0   'False
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
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ Ã” ÃÊÌ ﬂ«·«Â«"
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
            Left            =   3480
            TabIndex        =   19
            Top             =   120
            Width           =   2895
         End
      End
      Begin VB.Frame Frame53 
         Height          =   3975
         Left            =   -74640
         TabIndex        =   3
         Top             =   4080
         Width           =   6615
         Begin VB.CheckBox chkMultiInventory 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5760
            TabIndex        =   224
            Top             =   3480
            Width           =   375
         End
         Begin VB.TextBox txtReportHeaderName 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   360
            TabIndex        =   217
            Text            =   "¬—Ì«"
            Top             =   3000
            Width           =   2415
         End
         Begin VB.ComboBox CmbCycleStockNoDefault 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   360
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2450
            Width           =   2475
         End
         Begin VB.ComboBox cmbServePlace 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   360
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   155
            Width           =   2475
         End
         Begin VB.ComboBox cmbPartition 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   360
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   610
            Width           =   2475
         End
         Begin VB.ComboBox cboCustFind 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmStationsetting.frx":A59C
            Left            =   360
            List            =   "frmStationsetting.frx":A5AC
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1085
            Width           =   2475
         End
         Begin VB.ComboBox CmbFactorSort 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   360
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1550
            Width           =   2475
         End
         Begin VB.ComboBox cmbInventory 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   360
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2000
            Width           =   2475
         End
         Begin VB.Label Label84 
            Alignment       =   1  'Right Justify
            Caption         =   "«‰ Œ«» ç‰œ «‰»«— œ— Ìò ›«ò Ê— ›—Ê‘"
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
            Left            =   1560
            TabIndex        =   225
            Top             =   3480
            Width           =   3975
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "‰«„ „Õ· «” ›«œÂ œ— ç«Å"
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
            Left            =   3480
            TabIndex        =   218
            Top             =   3000
            Width           =   2655
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ «‰»«— »—«Ì ‘„«—‘ Ê «‰ ﬁ«·"
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
            Left            =   3000
            TabIndex        =   15
            Top             =   2535
            Width           =   3135
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ Õ«·  ›—Ê‘"
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
            Left            =   3960
            TabIndex        =   13
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ »Œ‘"
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
            Left            =   3720
            TabIndex        =   12
            Top             =   627
            Width           =   2295
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ Ã” ÃÊÌ „‘ —òÌ‰"
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
            Left            =   3360
            TabIndex        =   11
            Top             =   1014
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷  — Ì» ﬂ«·«Â« œ— ç«Å"
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
            Left            =   3480
            TabIndex        =   10
            Top             =   1521
            Width           =   2535
         End
         Begin VB.Label Label70 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÌ‘ ›—÷ «‰»«— »—«Ì Œ—Ìœ"
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
            Left            =   3600
            TabIndex        =   9
            Top             =   2028
            Width           =   2415
         End
      End
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„œÌ—Ì  Ê  ‰ŸÌ„«  «Ì” ê«Â"
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
      Left            =   5250
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   -30
      Width           =   4095
   End
End
Attribute VB_Name = "frmStationsetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Parameter() As Parameter
Dim i, TmpCustomerServeplace, tmpSearch As Integer, tmpSearchOrder As Integer
Dim tmpStationNo As Integer
Dim IsFileExist As Boolean
Dim filetemp As New FileSystemObject
Dim tempstring As TextStream
Dim rctmp As New ADODB.Recordset

Private Sub ChkFixRateChange_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If ChkFixRateChange.Value = 1 Then
    ChkShiftRate.Value = 0
End If
End Sub

Private Sub ChkShiftRate_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If ChkShiftRate.Value = 1 Then
    ChkFixRateChange.Value = 0
End If
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
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
Private Sub cboStations_Click()
    tmpStationNo = cboStations.ItemData(cboStations.ListIndex)
'    MyFormAddEditMode = ViewKey
    StationSettingFile = App.Path & "\Setting\Station" & tmpStationNo & ".txt"
    IsFileExist = filetemp.FileExists(StationSettingFile)
    
    If IsFileExist = False Then
      SetDefaultStationSettingFile
      MsgBox "Station Setting File Did Not Exist" & vbCrLf & "Default Station Setting File Created"
    
    End If
    Call clsStation.Class_Initialize
    DefaultSetting

End Sub
Private Sub FillStationCombo()
    On Error GoTo Err_Handler
    
    Dim Rst As New ADODB.Recordset
    Set Rst = RunStoredProcedure2RecordSet("Get_Pc_Stations")

    cboStations.Clear
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Not Rst.EOF
            cboStations.AddItem Rst.Fields("Description").Value
            cboStations.ItemData(cboStations.ListCount - 1) = Rst.Fields("StationID").Value
            Rst.MoveNext
        Wend
    End If
    
    For i = 0 To cboStations.ListCount - 1
        'cboStations.ListIndex = i
        If clsArya.StationNo = cboStations.ItemData(i) Then
            cboStations.ListIndex = i
            Exit For
        End If
    Next
    If cboStations.ListIndex <> -1 Then tmpStationNo = cboStations.ItemData(cboStations.ListIndex)
    
    Rst.Close
    Exit Sub
    
Err_Handler:
    LogSaveNew "FrmStationSetting => ", err.Description, err.Number, err.Source, "FillStationCombo"
    ShowErrorMessage
    err.Clear
End Sub

Private Sub Form_Load()
    
  ' Label10.Caption = "„œÌ—Ì  Ê  ‰ŸÌ„«  «Ì” ê«Â ‘„«—Â " & clsArya.StationNo
    
    CenterTop Me
        
    SetFirstToolBar
    
'    FillCmbNotice
    FillStationCombo
    

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

    formloadFlag = True
    
End Sub
Private Sub DefaultSetting()
    On Error GoTo Err_Handler
   
   Label10.Caption = "„œÌ—Ì  Ê  ‰ŸÌ„«  «Ì” ê«Â ‘„«—Â "  '& clsArya.StationNo
    
    Dim rctmp As New ADODB.Recordset

    If intVersion = Min Then
        Frame_Rate.Enabled = False
        Frame_CallerId.Enabled = False
        Frame_Control.Enabled = False
        Frame_Kala.Enabled = False
    End If
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_Serveplace", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            cmbServePlace.AddItem IIf(IsNull(rctmp.Fields("Description").Value), "", rctmp.Fields("Description").Value)
            cmbServePlace.ItemData(cmbServePlace.ListCount - 1) = rctmp.Fields("intServePlace").Value
            rctmp.MoveNext
        Wend
        For i = 0 To cmbServePlace.ListCount - 1
            If cmbServePlace.ItemData(i) = clsStation.ServePlaceDefault Then
               Me.cmbServePlace.ListIndex = i
            End If
        Next i
    End If
    rctmp.Close
    ''''''
    cmbInventory.Clear
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Type", adInteger, 4, 0) 'Central Inventory
    Set rctmp = RunParametricStoredProcedure2Rec("GetInventory", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
            cmbInventory.AddItem rctmp.Fields("Description")
            cmbInventory.ItemData(cmbInventory.ListCount - 1) = Val(rctmp.Fields("InventoryNo"))
            rctmp.MoveNext
        Loop
        For i = 1 To cmbInventory.ListCount
            If cmbInventory.ItemData(i - 1) = clsStation.PurchaseInventoryDefault Then
               Me.cmbInventory.ListIndex = i - 1
               Exit For
            End If
        Next i
        
    End If
    rctmp.Close
    
    ''''''
    CmbCycleStockNoDefault.Clear
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Type", adInteger, 4, 0) 'Central Inventory
    Set rctmp = RunParametricStoredProcedure2Rec("GetInventory", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
            CmbCycleStockNoDefault.AddItem rctmp.Fields("Description")
            CmbCycleStockNoDefault.ItemData(CmbCycleStockNoDefault.ListCount - 1) = Val(rctmp.Fields("InventoryNo"))
            rctmp.MoveNext
        Loop
        For i = 1 To CmbCycleStockNoDefault.ListCount
            If CmbCycleStockNoDefault.ItemData(i - 1) = clsStation.CycleStockNoDefault Then
               CmbCycleStockNoDefault.ListIndex = i - 1
               Exit For
            End If
        Next i
        
    End If
    rctmp.Close
    
    ''''''
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_All_tPartitions", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            cmbPartition.AddItem rctmp.Fields("PartitionDescription").Value
            cmbPartition.ItemData(cmbPartition.ListCount - 1) = rctmp.Fields("PartitionID").Value
            rctmp.MoveNext
        Wend
        For i = 0 To cmbPartition.ListCount - 1
            If cmbPartition.ItemData(i) = clsStation.PartitionID Then
               Me.cmbPartition.ListIndex = i
            End If
        Next i
    End If
    rctmp.Close
     
    cmbTypeBascule.AddItem "”—Ì Å‰œ-TSC"
    cmbTypeBascule.ItemData(0) = EnumTypeBascule.Pand
    cmbTypeBascule.AddItem "”—Ì œÌ ÃÌ"
    cmbTypeBascule.ItemData(1) = EnumTypeBascule.Digi
    cmbTypeBascule.AddItem "”—Ì Å‰œ-TLP"
    cmbTypeBascule.ItemData(2) = EnumTypeBascule.Pand_TLP
    For i = 0 To cmbTypeBascule.ListCount - 1
        If cmbTypeBascule.ItemData(i) = clsStation.TypeBascule Then
           Me.cmbTypeBascule.ListIndex = i
           Exit For
        End If
    Next i
    CmbCustomerRate.Clear
    CmbCustomerRate.AddItem " ›Ì ›—Ê‘ 1"
    CmbCustomerRate.ItemData(0) = 0
    CmbCustomerRate.AddItem " ›Ì ›—Ê‘ 2"
    CmbCustomerRate.ItemData(1) = 1
    CmbCustomerRate.AddItem " ›Ì ›—Ê‘ 3"
    CmbCustomerRate.ItemData(2) = 2
    CmbCustomerRate.AddItem " ›Ì ›—Ê‘ 4"
    CmbCustomerRate.ItemData(3) = 3
    CmbCustomerRate.AddItem " ›Ì ›—Ê‘ 5"
    CmbCustomerRate.ItemData(4) = 4
    CmbCustomerRate.AddItem " ›Ì ›—Ê‘ 6"
    CmbCustomerRate.ItemData(5) = 5
    
    For i = 0 To CmbCustomerRate.ListCount - 1
        If CmbCustomerRate.ItemData(i) = clsStation.CustomerRate Then
           Me.CmbCustomerRate.ListIndex = i
           Exit For
        End If
    Next i
    CmbFromStoreFee.Clear
    CmbFromStoreFee.AddItem " „Ì«‰êÌ‰ Œ—Ìœ"
    CmbFromStoreFee.ItemData(0) = 0
    CmbFromStoreFee.AddItem " ›Ì Œ—Ìœ"
    CmbFromStoreFee.ItemData(1) = 1
    CmbFromStoreFee.AddItem " ›Ì ›—Ê‘ 1"
    CmbFromStoreFee.ItemData(2) = 2
    CmbFromStoreFee.AddItem " ›Ì ›—Ê‘ 2"
    CmbFromStoreFee.ItemData(3) = 3
    CmbFromStoreFee.AddItem " ›Ì ›—Ê‘ 3"
    CmbFromStoreFee.ItemData(4) = 4
    CmbFromStoreFee.AddItem " ›Ì ›—Ê‘ 4"
    CmbFromStoreFee.ItemData(5) = 5
    CmbFromStoreFee.AddItem " ›Ì ›—Ê‘ 5"
    CmbFromStoreFee.ItemData(6) = 6
    CmbFromStoreFee.AddItem " ›Ì ›—Ê‘ 6"
    CmbFromStoreFee.ItemData(7) = 7
    Me.CmbFromStoreFee.ListIndex = 0
    For i = 0 To CmbFromStoreFee.ListCount - 1
        If CmbFromStoreFee.ItemData(i) = clsStation.FromStoreFee Then
           Me.CmbFromStoreFee.ListIndex = i
           Exit For
        End If
    Next i
   
    
    CmbFactorSort.AddItem "òœò«·«"
    CmbFactorSort.ItemData(0) = 0
    CmbFactorSort.AddItem "Õ—Ê› «·›»«"
    CmbFactorSort.ItemData(1) = 1
    CmbFactorSort.AddItem "ﬁÌ„ "
    CmbFactorSort.ItemData(2) = 2
    CmbFactorSort.AddItem " — Ì» Ê—Êœ"
    CmbFactorSort.ItemData(3) = 3
    For i = 0 To CmbFactorSort.ListCount - 1
        If CmbFactorSort.ItemData(i) = clsStation.FactorSortItems Then
           Me.CmbFactorSort.ListIndex = i
        End If
    Next i
    txtBarcodePrice.Text = clsStation.PriceChance
    
    If clsStation.DeliveryBarcodeDefault = 0 Then
        OptDelivery(0).Value = True
    Else
        OptDelivery(1).Value = True
    End If
    If clsStation.TableBarcodeDefault = 0 Then
        OptTable(0).Value = True
    Else
        OptTable(1).Value = True
    End If
    If clsStation.ReprintDefault = 0 Then
        OptPrint(0).Value = True
    Else
        OptPrint(1).Value = True
    End If
 
    If clsStation.DeliveryNoView = True Then
       ChDelView.Value = 1
    Else
       ChDelView.Value = 0
    End If
    If clsStation.AutoDrawerOpen = True Then
       ChOpenDrawer.Value = 1
    Else
       ChOpenDrawer.Value = 0
    End If
    
    If clsStation.ChangeGoodPrint = True Then
       ChChangeGoodPrint.Value = 1
    Else
       ChChangeGoodPrint.Value = 0
    End If
      
    If clsStation.AlphabeticGoods = True Then
       ChAlphabetic.Value = 1
    Else
       ChAlphabetic.Value = 0
    End If
    If clsStation.TableControl = True Then
       CheckTable.Value = 1
    Else
       CheckTable.Value = 0
    End If
    Set rctmp = Nothing
    
    
    If clsStation.StopOnEditFich = True Then
       ChkStopOnEditFich.Value = 1
    Else
       ChkStopOnEditFich.Value = 0
    End If
    
    If clsStation.SearchType = 0 Then
        OptSearch(0).Value = True
    ElseIf clsStation.SearchType = 1 Then
        OptSearch(1).Value = True
    ElseIf clsStation.SearchType = 2 Then
        OptSearch(2).Value = True
    ElseIf clsStation.SearchType = 3 Then
        OptSearch(3).Value = True
    End If
    
    If clsStation.DeletedGood = True Then
        OptDeletedGood(0).Value = True
    Else
        OptDeletedGood(1).Value = True
    End If
    
    If clsStation.SearchFichDefault = True Then
        OptSearchFichDefault(0).Value = True
    Else
        OptSearchFichDefault(1).Value = True
    End If
    
    If clsStation.CustomerOrderDefault = True Then
        OptcustOrder(0).Value = True
    Else
        OptcustOrder(1).Value = True
    End If
    
    If clsStation.CustomerSearchDefault = True Then
        OptCustSearch(0).Value = True
    Else
        OptCustSearch(1).Value = True
    End If
   
    If clsStation.CreditCalculate = True Then
       ChkCredit.Value = 1
    Else
       ChkCredit.Value = 0
    End If
   
    If clsStation.GoodSearchDefault = True Then
        OptGoodSearch(0).Value = True
    Else
        OptGoodSearch(1).Value = True
    End If
   
    If clsStation.MojodiControlDefault = True Then
        ChkMojodiControlDefault.Value = 1
    Else
        ChkMojodiControlDefault.Value = 0
    End If
    If clsStation.RowMojodiControl = True Then
        ChkRowMojodiControl.Value = 1
    Else
        ChkRowMojodiControl.Value = 0
    End If
    
    If clsStation.RefreshFichNo = True Then
        ChkRefreshFichNo.Value = 1
    Else
        ChkRefreshFichNo.Value = 0
    End If
    
    If clsStation.DirectBascule = True Then
        ChkDirectBascule.Value = 1
    Else
        ChkDirectBascule.Value = 0
    End If
    
    If clsStation.BarcodeChance = True Then
        ChkBarcodeChance.Value = 1
    Else
        ChkBarcodeChance.Value = 0
    End If
    
'    If clsStation.CommandView = True Then
'        ChkCommandView.Value = 1
'    Else
'        ChkCommandView.Value = 0
'    End If
'
    If clsStation.DiscountDefault = 0 Then
        OptDiscount(0).Value = True
    Else
        OptDiscount(1).Value = True
    End If
    
    If clsStation.StartUpFormDefault = 0 Then
        OptStartForm(0).Value = True
    ElseIf clsStation.StartUpFormDefault = 1 Then
        OptStartForm(1).Value = True
    Else
        OptStartForm(2).Value = True
    End If
   
   
   If clsStation.CustomerServeplace = 0 Then
        OptcustServePlace(0).Value = True
    ElseIf clsStation.CustomerServeplace = 1 Then
        OptcustServePlace(1).Value = True
    ElseIf clsStation.CustomerServeplace = 2 Then
        OptcustServePlace(2).Value = True
    End If
    
    If clsStation.MultiPrice = True And clsArya.MultiPrice = True Then
        ChkMultiPrice.Value = 1
    Else
        ChkMultiPrice.Value = 0
    End If
     
    If clsStation.ShiftRate = True Then
        ChkShiftRate.Value = 1
    Else
        ChkShiftRate.Value = 0
    End If
     
    
    If clsStation.UpdateBuyPrice = True Then
        ChkUpdateBuyPrice.Value = 1
    Else
        ChkUpdateBuyPrice.Value = 0
    End If
    
    If clsStation.UpdateSellprice = True Then
        ChkUpdateSellPrice.Value = 1
    Else
        ChkUpdateSellPrice.Value = 0
    End If
    
    If clsStation.TrazooBarcode = True Then
        ChkTrazooBarcode.Value = 1
    Else
        ChkTrazooBarcode.Value = 0
    End If
''    If clsStation.CodeFlag = True Then
''        ChkCodeFlag.Value = 1
''    Else
''        ChkCodeFlag.Value = 0
''    End If
    
    If clsStation.SoundAlarm = True Then
        ChkSoundAlarm.Value = 1
    Else
        ChkSoundAlarm.Value = 0
    End If
    
    If clsStation.CashPayment = True Then
        chkCashPayment.Value = 1
    Else
        chkCashPayment.Value = 0
    End If
    
    If clsStation.InpersonSalonPayment = True Then
        ChkInpersonSalonPayment.Value = 1
    Else
        ChkInpersonSalonPayment.Value = 0
    End If
    
    If clsStation.InpersonDeliveryPayment = True Then
        ChkInpersonDeliveryPayment.Value = 1
    Else
        ChkInpersonDeliveryPayment.Value = 0
    End If
    
    If clsStation.InpersonOutPayment = True Then
        ChkInpersonOutPayment.Value = 1
    Else
        ChkInpersonOutPayment.Value = 0
    End If
    
    If clsStation.InpersonTablePayment = True Then
        ChkInpersonTablePayment.Value = 1
    Else
        ChkInpersonTablePayment.Value = 0
    End If
    
    If clsStation.InpersonSalonBalance = True Then
        ChkInpersonSalonBalance.Value = 1
    Else
        ChkInpersonSalonBalance.Value = 0
    End If
    
    If clsStation.InpersonDeliveryBalance = True Then
        ChkInpersonDeliveryBalance.Value = 1
    Else
        ChkInpersonDeliveryBalance.Value = 0
    End If
    
    If clsStation.InpersonOutBalance = True Then
        ChkInpersonOutBalance.Value = 1
    Else
        ChkInpersonOutBalance.Value = 0
    End If
    
    If clsStation.InpersonTableBalance = True Then
        ChkInpersonTableBalance.Value = 1
    Else
        ChkInpersonTableBalance.Value = 0
    End If
    
    If clsStation.ByPhoneSalonPayment = True Then
        ChkByPhoneSalonPayment.Value = 1
    Else
        ChkByPhoneSalonPayment.Value = 0
    End If
    
    If clsStation.ByPhoneDeliveryPayment = True Then
        ChkByPhoneDeliveryPayment.Value = 1
    Else
        ChkByPhoneDeliveryPayment.Value = 0
    End If
    
    If clsStation.ByPhoneTablePayment = True Then
        ChkByPhoneTablePayment.Value = 1
    Else
        ChkByPhoneTablePayment.Value = 0
    End If
    
    If clsStation.ByPhoneSalonBalance = True Then
        ChkByPhoneSalonBalance.Value = 1
    Else
        ChkByPhoneSalonBalance.Value = 0
    End If
    
    If clsStation.ByPhoneDeliveryBalance = True Then
        ChkByPhoneDeliveryBalance.Value = 1
    Else
        ChkByPhoneDeliveryBalance.Value = 0
    End If
    
    If clsStation.ByPhoneTableBalance = True Then
        ChkByPhoneTableBalance.Value = 1
    Else
        ChkByPhoneTableBalance.Value = 0
    End If
    
    If clsStation.ThreeSegmentSearch = True Then
        ChkThreeSegmentSearch.Value = 1
    Else
        ChkThreeSegmentSearch.Value = 0
    End If
    
    If clsStation.NumberOfUnitSale = True Then
        ChkNumberOfUnit.Value = 1
    Else
        ChkNumberOfUnit.Value = 0
    End If
    
    If clsStation.MenuViewAfterGood = True Then
        ChkMenuViewAfterGood.Value = 1
    Else
        ChkMenuViewAfterGood.Value = 0
    End If
    
    If clsStation.PayFactorView = True Then
        ChkPayFactorView.Value = 1
    Else
        ChkPayFactorView.Value = 0
    End If
    
    If clsStation.BarcodeAutoEscape = True Then
        ChkBarcodeAutoEscape.Value = 1
    Else
        ChkBarcodeAutoEscape.Value = 0
    End If
    If clsStation.PrintAfterTasvieh = True Then
        ChkPrintAfterTasvieh.Value = 1
    Else
        ChkPrintAfterTasvieh.Value = 0
    End If
    If clsStation.AutoBarcode = True Then
        ChkAutoBarcode.Value = 1
    Else
        ChkAutoBarcode.Value = 0
    End If
    
    If clsStation.AutoCallerId = True Then
        ChkAutoCallerId.Value = 1
    Else
        ChkAutoCallerId.Value = 0
    End If
    
    If clsStation.CustomerAscii = True Then
        ChkCustomerAscii.Value = 1
    Else
        ChkCustomerAscii.Value = 0
    End If
    
    If clsStation.CustomerFarsi = True Then
        ChkCustomerFarsi.Value = 1
    Else
        ChkCustomerFarsi.Value = 0
    End If
    
    If clsStation.CustomerOnlinePrice = True Then
        ChkCustomerOnlinePrice.Value = 1
    Else
        ChkCustomerOnlinePrice.Value = 0
    End If
    
    If clsStation.CustomerFarsi = True Then
        ChkCustomerFarsi.Value = 1
    Else
        ChkCustomerFarsi.Value = 0
    End If
    
    If clsStation.NoCurrentDay = True Then
        ChkNoCurrentDay.Value = 1
    Else
        ChkNoCurrentDay.Value = 0
    End If
    
'''    If clsStation.Callwaiting = True Then
'''        ChkCallwaiting.Value = 1
'''    Else
'''        ChkCallwaiting.Value = 0
'''    End If
    
     If clsStation.SearchOrderType = 0 Then
        OptOrder(0).Value = True
    ElseIf clsStation.SearchOrderType = 1 Then
        OptOrder(1).Value = True
    ElseIf clsStation.SearchOrderType = 2 Then
        OptOrder(2).Value = True
    End If
    If clsStation.PriceType = 0 Then clsStation.PriceType = 1
    If clsStation.MaxPrices = 0 Then clsStation.MaxPrices = 1
    If clsStation.OutPrice = 0 Then clsStation.OutPrice = 1
    
    FWNumericPrice.Value = clsStation.PriceType
    FWMaxNumericPrice.Value = clsStation.MaxPrices
    FWNumericOutPrice.Value = clsStation.OutPrice
    
    cboCustFind.ListIndex = clsStation.DefaultCustSearch
    
    TxtCallerIdSpace.Text = Val(clsStation.CallerIdSpace)
    TxtCountCustomerDailyBuy.Text = Val(clsStation.CountCustomerDailyBuy)
    txtCountCustomerGoods.Text = Val(clsStation.CountCustomerGood)
    TxtInvoiceRows.Text = Val(clsStation.InvoiceRows)
    TxtPurchaseRows.Text = Val(clsStation.PurchaseRows)
    If clsStation.FinalCheck = True Then
       ChkFinalCheck.Value = 1
    Else
       ChkFinalCheck.Value = 0
    End If
    
    If clsStation.AutoTip = True Then
       ChkAutoTip.Value = 1
    Else
       ChkAutoTip.Value = 0
    End If
     
'    If clsStation.FichStatusBar = True Then
'       ChkFichStatusBar.Value = 1
'    Else
'       ChkFichStatusBar.Value = 0
'    End If
    
     If clsStation.TextIconViewH = True Then
       ChkTextIconViewH.Value = 1
    Else
       ChkTextIconViewH.Value = 0
    End If
    
'     If clsStation.TextIconViewv = True Then
'       ChkTextIconViewV.Value = 1
'    Else
'       ChkTextIconViewV.Value = 0
'    End If
    
'    If clsStation.ForceSeller = True Then
'        ChkForceSeller.Value = 1
'    Else
'        ChkForceSeller.Value = 0
'    End If
    
    If clsStation.InvoiceStatusDefault = True Then
        ChkInvoiceStatusDefault.Value = 1
    Else
        ChkInvoiceStatusDefault.Value = 0
    End If
    
    TxtCityCode.Text = IIf(Val(clsStation.CityCode) = 0, "21", Val(clsStation.CityCode))
    TxtNumberOfId.Text = IIf(Val(clsStation.NumberOfId) = 0, "8", Val(clsStation.NumberOfId))
    
    TxtDiscoveryPort.Text = IIf(Val(clsStation.DiscoveryPort) = 0, "4111", Val(clsStation.DiscoveryPort))
    TxtResponsePort.Text = IIf(Val(clsStation.ResponsePort) = 0, "4112", Val(clsStation.ResponsePort))
    
    If clsStation.AlmLogFile = True Then
        ChkAlmLogFile.Value = 1
    Else
        ChkAlmLogFile.Value = 0
    End If
    
    If clsStation.NetworkCallerId = True Then
        ChkNetworkCallerId.Value = 1
    Else
        ChkNetworkCallerId.Value = 0
    End If
    
    If clsStation.CustomerFeeDataBase = True Then
        ChkCustomerFeeDatabase.Value = 1
    Else
        ChkCustomerFeeDatabase.Value = 0
    End If
    
    TxtStartNumberCartReader.Text = IIf(Val(clsStation.StartNumberCartReader) = 0, "0", Val(clsStation.StartNumberCartReader))
    TxtNumberOfCardReader.Text = IIf(Val(clsStation.NumberOfCardReader) = 0, "10", Val(clsStation.NumberOfCardReader))
   
      
'    If clsStation.TaxView = True Then
'        ChkTaxView.Value = 1
'    Else
'        ChkTaxView.Value = 0
'    End If
    
    If clsStation.EditCompatibleSamar1 = True Then
        ChkEditCompatibleSamar1.Value = 1
    Else
        ChkEditCompatibleSamar1.Value = 0
    End If
    
    If clsStation.UndoRedoCompatibleSamar1 = True Then
        ChkUndoRedoCompatibleSamar1.Value = 1
    Else
        ChkUndoRedoCompatibleSamar1.Value = 0
    End If
    
    If clsStation.ViewTempAddress = True Then
        OptViewTempAddress(0).Value = True
    Else
        OptViewTempAddress(1).Value = True
    End If
    
    If clsStation.AlphabetGoodSearch = True Then
        OptAlphabetGoodSearch(0).Value = True
    Else
        OptAlphabetGoodSearch(1).Value = True
    End If
    
    TxtCountCustomerShiftBuy.Text = Val(clsStation.CountCustomerShiftBuy)
    
    If clsStation.PrintAfterDeliver = True Then
        ChkPrintAfterDeliver.Value = 1
    Else
        ChkPrintAfterDeliver.Value = 0
    End If
    
    If clsStation.FixRateChange = True Then
        ChkFixRateChange.Value = 1
    Else
        ChkFixRateChange.Value = 0
    End If
    
    If clsStation.AutoCashClose = True Then
       ChkAutoCashClose.Value = 1
    Else
       ChkAutoCashClose.Value = 0
    End If
    
    If clsStation.UpDateCarryFee = True Then
        ChkUpDateCarryFee.Value = 1
    Else
        ChkUpDateCarryFee.Value = 0
    End If

    If clsStation.CallerIdAutoView = True Then
        ChkfrmCallerId.Value = 1
    Else
        ChkfrmCallerId.Value = 0
    End If

    If clsStation.AutoBackup = True Then
        ChkAutoBackUp.Value = 1
    Else
        ChkAutoBackUp.Value = 0
    End If
    If clsStation.PosPayment = True Then
        ChkPosPayment.Value = 1
    Else
        ChkPosPayment.Value = 0
    End If

'    TxtPassPhrase.Text = clsStation.PassPhrase
'    TxtPosApprovedText.Text = clsStation.PosApprovedText
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblPub_Pos_ByStationId", Parameter)
    
    CmbPosModel.Clear
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Not Rst.EOF
            i = i + 1
            CmbPosModel.AddItem Rst.Fields("nvcBankName").Value
            CmbPosModel.ItemData(CmbPosModel.ListCount - 1) = Rst.Fields("PosId").Value
            Rst.MoveNext
        Wend
    End If
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
    For i = 0 To CmbPosModel.ListCount - 1
        If CmbPosModel.ItemData(i) = clsStation.PosModel Then
           Me.CmbPosModel.ListIndex = i
           Exit For
        End If
    Next i
    If clsStation.ReportHeadername <> "" Then
        txtReportHeaderName.Text = clsStation.ReportHeadername
    End If
    
    If clsStation.FastCustSave = True Then
        ChkFastCustSave.Value = 1
    Else
        ChkFastCustSave.Value = 0
    End If
    
    If clsStation.RepetitiveGood = True Then
       ChkRepetitiveGood.Value = 1
    Else
       ChkRepetitiveGood.Value = 0
    End If

    If clsStation.OtherPartition = True Then
       chkOtherPartition.Value = 1
    Else
       chkOtherPartition.Value = 0
    End If

    If clsStation.NotShowPrintNotice = True Then
       chkNotShowPrintNotice.Value = 1
    Else
       chkNotShowPrintNotice.Value = 0
    End If

    If clsStation.MultiInventory = True Then
       chkMultiInventory.Value = 1
    Else
       chkMultiInventory.Value = 0
    End If

    If clsStation.Frame_Printers = True Then
       chkFrame_Printers.Value = 1
    Else
       chkFrame_Printers.Value = 0
    End If

    If clsStation.HasOptionPrice = True Then
       chkHasOptionPrice.Value = 1
    Else
       chkHasOptionPrice.Value = 0
    End If

    If clsStation.ShowOption = True Then
       chkShowOption.Value = 1
    Else
       chkShowOption.Value = 0
    End If

    If clsStation.LableUsedGood = True Then
       chkLableUsedGood.Value = 1
    Else
       chkLableUsedGood.Value = 0
    End If

    If clsStation.LoyaltyCustomers = True Then
       chkLoyaltyCustomers.Value = 1
    Else
       chkLoyaltyCustomers.Value = 0
    End If

    If clsStation.LoyaltyAllCustomers = True Then
       chkLoyaltyAllCustomers.Value = 1
    Else
       chkLoyaltyAllCustomers.Value = 0
    End If

    txtStartCharacter.Text = clsStation.StartCharacter

    If clsStation.OneClickShow = True Then
       chkOneClickShow.Value = 1
    Else
       chkOneClickShow.Value = 0
    End If

    If clsStation.TouchScreen = True Then
       ChkTouchScreen.Value = 1
    Else
       ChkTouchScreen.Value = 0
    End If

    FWNoRowMenu.Value = clsStation.NoRowMenu

    If clsStation.RfidReader = True Then
        chkRfidReader.Value = 1
    Else
        chkRfidReader.Value = 0
    End If

    txtRfidInterval = IIf(Val(clsStation.RfidInterval) = 0, "2000", Val(clsStation.RfidInterval))

    If clsStation.RfidLongBuzzer = True Then
        ChkRfidLongBeep.Value = 1
    Else
        ChkRfidLongBeep.Value = 0
    End If

    If clsStation.TelNetServerActive = True Then
        chkTelNetServerActive.Value = 1
    Else
        chkTelNetServerActive.Value = 0
    End If

    txtTelnetServerIP = IIf(clsStation.TelNetServerIP = "", "192.168.1.1", clsStation.TelNetServerIP)

    txtTelnetServerPort = IIf(clsStation.TelNetServerPort = "", "2001", clsStation.TelNetServerPort)
    If clsStation.PrintAfterPayk = True Then
        ChkPrintAfterPayk.Value = 1
    Else
        ChkPrintAfterPayk.Value = 0
    End If

    If clsStation.PrintAfterOrder = True Then
        chkPrintAfterOrder.Value = 1
    Else
        chkPrintAfterOrder.Value = 0
    End If
    
    If clsStation.LabelPrint = True Then
        chkLabelPrint.Value = 1
    Else
        chkLabelPrint.Value = 0
    End If
    
    If clsStation.AryaSmsPanel = True Then
        chkAryaSmsPanel.Value = 1
    Else
        chkAryaSmsPanel.Value = 0
    End If
    
    If clsStation.ForceTax = True Then
        chkForceTax.Value = 1
    Else
        chkForceTax.Value = 0
    End If
    
    If clsStation.PersonIdCheck = True Then
       chkPersonIdCheck.Value = 1
    Else
       chkPersonIdCheck.Value = 0
    End If

    txtPersonIdRefreshTime = IIf(Val(clsStation.PersonIdRefreshTime) = 0, "5", Val(clsStation.PersonIdRefreshTime))

    If clsStation.ListAutoLoad = True Then
       chkListAutoLoad.Value = 1
    Else
       chkListAutoLoad.Value = 0
    End If
    txtDeviceIP = IIf(clsStation.DeviceIP = "", "192.168.1.100", clsStation.DeviceIP)
    txtDeviceId = Val(clsStation.DeviceID)
    txtListFont = IIf(Val(clsStation.ListFont) = 0, "14", Val(clsStation.ListFont))

    txtDevice2IP = IIf(clsStation.Device2IP = "", "192.168.1.101", clsStation.Device2IP)
    txtDevice2Id = Val(clsStation.Device2Id)

Exit Sub
Err_Handler:
    ShowDisMessage err.Description, 3000

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    VarActForm = ""

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


End Sub

Public Sub Update()

    If cmbServePlace.ListIndex <> -1 Then
        clsStation.ServePlaceDefault = cmbServePlace.ItemData(cmbServePlace.ListIndex)
    End If
    If cmbInventory.ListIndex <> -1 Then
        clsStation.PurchaseInventoryDefault = cmbInventory.ItemData(cmbInventory.ListIndex)
    End If

    If CmbCycleStockNoDefault.ListIndex <> -1 Then
        clsStation.CycleStockNoDefault = CmbCycleStockNoDefault.ItemData(CmbCycleStockNoDefault.ListIndex)
    End If

    If cmbPartition.ListIndex <> -1 Then
        clsStation.PartitionID = cmbPartition.ItemData(cmbPartition.ListIndex)
    End If
    
    If cmbTypeBascule.ListIndex <> -1 Then
        clsStation.TypeBascule = cmbTypeBascule.ItemData(cmbTypeBascule.ListIndex)
    Else
        clsStation.TypeBascule = 0
    End If
    
    If CmbFromStoreFee.ListIndex <> -1 Then
        clsStation.FromStoreFee = CmbFromStoreFee.ItemData(CmbFromStoreFee.ListIndex)
    End If
    
    If CmbCustomerRate.ListIndex <> -1 Then
        clsStation.CustomerRate = CmbCustomerRate.ItemData(CmbCustomerRate.ListIndex)
    End If
    
    If OptcustServePlace(0).Value = True Then
        TmpCustomerServeplace = 0
    ElseIf OptcustServePlace(1).Value = True Then
        TmpCustomerServeplace = 1
    ElseIf OptcustServePlace(2).Value = True Then
        TmpCustomerServeplace = 2
    End If
    If OptSearch(0).Value = True Then
        tmpSearch = 0
    ElseIf OptSearch(1).Value = True Then
        tmpSearch = 1
    ElseIf OptSearch(2).Value = True Then
        tmpSearch = 2
    ElseIf OptSearch(3).Value = True Then
        tmpSearch = 3
    End If
    If OptOrder(0).Value = True Then
        tmpSearchOrder = 0
    ElseIf OptOrder(1).Value = True Then
        tmpSearchOrder = 1
    ElseIf OptOrder(2).Value = True Then
        tmpSearchOrder = 2
    End If
    Dim TempOptStartForm
    If OptStartForm(0).Value = True Then
        TempOptStartForm = 0
    ElseIf OptStartForm(1).Value = True Then
        TempOptStartForm = 1
    ElseIf OptStartForm(2).Value = True Then
        TempOptStartForm = 2
    End If
    clsStation.DefaultCustSearch = cboCustFind.ListIndex
    
    clsStation.PriceChance = Val(txtBarcodePrice.Text)
    clsStation.DeliveryBarcodeDefault = IIf(OptDelivery(0).Value = True, 0, 1)
    clsStation.TableBarcodeDefault = IIf(OptTable(0).Value = True, 0, 1)
    clsStation.ReprintDefault = IIf(OptPrint(0).Value = True, 0, 1)
    clsStation.DeliveryNoView = ChDelView.Value
    clsStation.AutoDrawerOpen = ChOpenDrawer.Value
    clsStation.ChangeGoodPrint = ChChangeGoodPrint.Value
    clsStation.AlphabeticGoods = ChAlphabetic.Value
    clsStation.TableControl = IIf(CheckTable.Value = 1, True, False)
    clsStation.StopOnEditFich = IIf(ChkStopOnEditFich.Value = 1, True, False)
    clsStation.SearchType = tmpSearch
    clsStation.SearchOrderType = tmpSearchOrder
    clsStation.DeletedGood = OptDeletedGood(0).Value
    clsStation.SearchFichDefault = OptSearchFichDefault(0).Value
    clsStation.CustomerOrderDefault = OptcustOrder(0).Value
    clsStation.CustomerServeplace = TmpCustomerServeplace
    clsStation.PriceType = FWNumericPrice.Value
    clsStation.MaxPrices = FWMaxNumericPrice.Value
    clsStation.CustomerSearchDefault = OptCustSearch(0).Value
    clsStation.GoodSearchDefault = OptGoodSearch(0).Value
    clsStation.CreditCalculate = ChkCredit.Value
    clsStation.DiscountDefault = IIf(OptDiscount(0).Value = True, 0, 1)
    clsStation.StartUpFormDefault = TempOptStartForm
    clsStation.MojodiControlDefault = ChkMojodiControlDefault.Value
    clsStation.RowMojodiControl = ChkRowMojodiControl.Value
    clsStation.RefreshFichNo = ChkRefreshFichNo.Value
    clsStation.DirectBascule = ChkDirectBascule.Value
    clsStation.BarcodeChance = ChkBarcodeChance.Value
    clsStation.FactorSortItems = CmbFactorSort.ListIndex
'    clsStation.CommandView = ChkCommandView.Value
    clsStation.MultiPrice = ChkMultiPrice.Value
    clsStation.UpdateBuyPrice = ChkUpdateBuyPrice.Value
    clsStation.UpdateSellprice = ChkUpdateSellPrice.Value
    clsStation.TrazooBarcode = ChkTrazooBarcode.Value
''    clsStation.CodeFlag = ChkCodeFlag.Value
    clsStation.SoundAlarm = ChkSoundAlarm.Value
    clsStation.CashPayment = chkCashPayment.Value
    
    clsStation.InpersonSalonPayment = ChkInpersonSalonPayment.Value
    clsStation.InpersonDeliveryPayment = ChkInpersonDeliveryPayment.Value
    clsStation.InpersonOutPayment = ChkInpersonOutPayment.Value
    clsStation.InpersonTablePayment = ChkInpersonTablePayment.Value
    
    clsStation.InpersonSalonBalance = ChkInpersonSalonBalance.Value
    clsStation.InpersonDeliveryBalance = ChkInpersonDeliveryBalance.Value
    clsStation.InpersonOutBalance = ChkInpersonOutBalance.Value
    clsStation.InpersonTableBalance = ChkInpersonTableBalance.Value
    
    clsStation.ByPhoneSalonPayment = ChkByPhoneSalonPayment.Value
    clsStation.ByPhoneDeliveryPayment = ChkByPhoneDeliveryPayment.Value
    clsStation.ByPhoneTablePayment = ChkByPhoneTablePayment.Value
    
    clsStation.ByPhoneSalonBalance = ChkByPhoneSalonBalance.Value
    clsStation.ByPhoneDeliveryBalance = ChkByPhoneDeliveryBalance.Value
    clsStation.ByPhoneTableBalance = ChkByPhoneTableBalance.Value
    clsStation.ThreeSegmentSearch = ChkThreeSegmentSearch.Value
    clsStation.NumberOfUnitSale = ChkNumberOfUnit.Value
    clsStation.MenuViewAfterGood = ChkMenuViewAfterGood.Value
    clsStation.PayFactorView = ChkPayFactorView.Value
    clsStation.BarcodeAutoEscape = ChkBarcodeAutoEscape
    clsStation.PrintAfterTasvieh = ChkPrintAfterTasvieh
    clsStation.AutoBarcode = ChkAutoBarcode
    clsStation.AutoCallerId = ChkAutoCallerId
    clsStation.CustomerAscii = ChkCustomerAscii
    clsStation.CustomerFarsi = ChkCustomerFarsi
    clsStation.CustomerOnlinePrice = ChkCustomerOnlinePrice
    clsStation.NoCurrentDay = ChkNoCurrentDay
    clsStation.CallerIdSpace = Val(TxtCallerIdSpace.Text)
''    clsStation.Callwaiting = ChkCallwaiting
    clsStation.CountCustomerDailyBuy = Val(TxtCountCustomerDailyBuy.Text)
    clsStation.CountCustomerGood = Val(txtCountCustomerGoods.Text)
'    clsStation.ForceSeller = ChkForceSeller
    clsStation.InvoiceRows = Val(TxtInvoiceRows.Text)
    clsStation.PurchaseRows = Val(TxtPurchaseRows.Text)
    clsStation.FinalCheck = IIf(ChkFinalCheck.Value = 1, True, False)
    clsStation.AutoTip = IIf(ChkAutoTip.Value = 1, True, False)
'    clsStation.FichStatusBar = IIf(ChkFichStatusBar.Value = 1, True, False)
    clsStation.TextIconViewH = IIf(ChkTextIconViewH.Value = 1, True, False)
'    clsStation.TextIconViewv = IIf(ChkTextIconViewV.Value = 1, True, False)
    clsStation.InvoiceStatusDefault = ChkInvoiceStatusDefault
    clsStation.CityCode = Val(TxtCityCode.Text)
    clsStation.NumberOfId = Val(TxtNumberOfId.Text)
    clsStation.DiscoveryPort = Val(TxtDiscoveryPort.Text)
    clsStation.ResponsePort = Val(TxtResponsePort.Text)
    clsStation.AlmLogFile = ChkAlmLogFile
    clsStation.NetworkCallerId = ChkNetworkCallerId
    clsStation.CustomerFeeDataBase = ChkCustomerFeeDatabase
    clsStation.StartNumberCartReader = Val(TxtStartNumberCartReader.Text)
    clsStation.NumberOfCardReader = Val(TxtNumberOfCardReader.Text)
'    clsStation.TaxView = ChkTaxView.Value
    clsStation.OutPrice = FWNumericOutPrice.Value
    clsStation.EditCompatibleSamar1 = ChkEditCompatibleSamar1
    clsStation.UndoRedoCompatibleSamar1 = ChkUndoRedoCompatibleSamar1
    clsStation.ViewTempAddress = OptViewTempAddress(0).Value
    clsStation.AlphabetGoodSearch = OptAlphabetGoodSearch(0).Value
    clsStation.ShiftRate = ChkShiftRate.Value
    clsStation.CountCustomerShiftBuy = Val(TxtCountCustomerShiftBuy.Text)
    clsStation.PrintAfterDeliver = ChkPrintAfterDeliver
    clsStation.FixRateChange = ChkFixRateChange.Value
    clsStation.AutoCashClose = IIf(ChkAutoCashClose.Value = 1, True, False)
    clsStation.UpDateCarryFee = ChkUpDateCarryFee
    clsStation.CallerIdAutoView = ChkfrmCallerId
    clsStation.AutoBackup = CBool(ChkAutoBackUp.Value)
    clsStation.PosPayment = ChkPosPayment
'    clsStation.PassPhrase = TxtPassPhrase.Text
'    clsStation.PosApprovedText = TxtPosApprovedText.Text
    clsStation.ReportHeadername = txtReportHeaderName.Text
    clsStation.FastCustSave = ChkFastCustSave.Value
    clsStation.RepetitiveGood = ChkRepetitiveGood.Value
    clsStation.OtherPartition = chkOtherPartition.Value
    clsStation.NotShowPrintNotice = chkNotShowPrintNotice.Value
    clsStation.MultiInventory = chkMultiInventory.Value
    clsStation.Frame_Printers = chkFrame_Printers.Value
    clsStation.HasOptionPrice = chkHasOptionPrice.Value
    clsStation.ShowOption = chkShowOption.Value
    clsStation.LableUsedGood = chkLableUsedGood.Value
    clsStation.LoyaltyCustomers = chkLoyaltyCustomers.Value
    clsStation.LoyaltyAllCustomers = chkLoyaltyAllCustomers.Value
    clsStation.StartCharacter = txtStartCharacter
    clsStation.OneClickShow = chkOneClickShow
    clsStation.TouchScreen = ChkTouchScreen
    clsStation.NoRowMenu = FWNoRowMenu.Value
    clsStation.RfidReader = chkRfidReader
    clsStation.RfidInterval = txtRfidInterval.Text
    clsStation.RfidLongBuzzer = ChkRfidLongBeep
    clsStation.TelNetServerActive = chkTelNetServerActive
    clsStation.TelNetServerIP = txtTelnetServerIP.Text
    clsStation.TelNetServerPort = txtTelnetServerPort.Text
    clsStation.PrintAfterPayk = ChkPrintAfterPayk.Value
    clsStation.PrintAfterOrder = chkPrintAfterOrder
    clsStation.LabelPrint = chkLabelPrint
    clsStation.AryaSmsPanel = chkAryaSmsPanel
    clsStation.ForceTax = chkForceTax
    clsStation.PersonIdCheck = chkPersonIdCheck.Value
    clsStation.PersonIdRefreshTime = Val(txtPersonIdRefreshTime.Text)
    clsStation.ListAutoLoad = chkListAutoLoad.Value
    clsStation.DeviceIP = txtDeviceIP.Text
    clsStation.DeviceID = Val(txtDeviceId.Text)
    clsStation.ListFont = Val(txtListFont.Text)
    clsStation.Device2IP = txtDevice2IP.Text
    clsStation.Device2Id = Val(txtDevice2Id.Text)
    
    If CmbPosModel.ListIndex <> -1 Then
        clsStation.PosModel = CmbPosModel.ItemData(CmbPosModel.ListIndex)
    End If
  
    StationSettingFile = App.Path & "\Setting\Station" & tmpStationNo & ".txt"
    IsFileExist = filetemp.FileExists(StationSettingFile)
    
    If IsFileExist = False Then
      SetDefaultStationSettingFile
      MsgBox "Station Setting File Did Not Exist" & vbCrLf & "Default Station Setting File Created"
    
    End If
    
    SetStationSettingFile
    
    ShowDisMessage "  ‰ŸÌ„«  «Ì” ê«Â ‘„«—Â " & tmpStationNo & " «‰Ã«„ ‘œ ", 1000
    
    Unload Me
    
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

Private Sub txtStartCharacter_Change()
    If Len(txtStartCharacter) > 1 Then txtStartCharacter = Right(txtStartCharacter, 1)
End Sub
