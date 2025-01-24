VERSION 5.00
Begin VB.Form frmFindGoods_Difference 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                  ÌÓÊÌæí ÂÔä åÇí ãäæ Çíä ÇíÓÊÇå             "
   ClientHeight    =   7275
   ClientLeft      =   6045
   ClientTop       =   2430
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFindGoods_Difference.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame_Option 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ÂÔä åÇí ˜ÇáÇ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   5295
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1920
      Width           =   6735
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   29
         Left            =   5520
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   28
         Left            =   4440
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   27
         Left            =   3360
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   26
         Left            =   2280
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   25
         Left            =   1200
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   24
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   23
         Left            =   5520
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   22
         Left            =   4440
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   21
         Left            =   3360
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   20
         Left            =   2280
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   19
         Left            =   1200
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   18
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   17
         Left            =   5520
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   16
         Left            =   4440
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   15
         Left            =   3360
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   14
         Left            =   2280
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   13
         Left            =   1200
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   12
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   11
         Left            =   5520
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   10
         Left            =   4440
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   9
         Left            =   3360
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   8
         Left            =   2280
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   7
         Left            =   1200
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   6
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   5
         Left            =   5520
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   4
         Left            =   4440
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   3
         Left            =   3360
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   2
         Left            =   2280
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   1
         Left            =   1200
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton BtnOption 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   0
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdEscape 
      BackColor       =   &H000000FF&
      Caption         =   "ÇäÕÑÇÝ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "Çß ßä"
         Height          =   1275
         Index           =   10
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   840
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Tag             =   "7"
         Top             =   960
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Tag             =   "8"
         Top             =   960
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
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Tag             =   "9"
         Top             =   960
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
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Tag             =   "4"
         Top             =   120
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Tag             =   "5"
         Top             =   960
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
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Tag             =   "6"
         Top             =   960
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
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "1"
         Top             =   120
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "2"
         Top             =   120
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "3"
         Top             =   120
         Width           =   795
      End
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
         Left            =   1080
         TabIndex        =   1
         Tag             =   "0"
         Top             =   120
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmFindGoods_Difference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumberCode As Integer
Dim i, TmpLineCode As Long
Dim TmpGoodName1, TmpGoodName2 As String
Dim j As Integer
Dim Parameter() As Parameter
Dim NotSupportedGoodType As EnumGoodType
Dim mvarGoodCode As Double
Dim CostOption(0 To 29) As Long
Public Function SendVariables(ByRef GoodCode)
    mvarGoodCode = GoodCode
End Function

Private Sub CancelButton_Click()
    If LCase(VarActForm) = "frminvoice" Then
        frmInvoice.lblNum.Caption = ""
    End If
    mvarcode = 0
    Unload Me
End Sub


Private Sub BtnKeypad_Click(index As Integer)
    If BtnKeypad(index).Tag = "" Then
        If LCase(VarActForm) = "frminvoice" Then
            If Len(Trim(frmInvoice.lblNum.Caption)) >= 1 Then
                frmInvoice.lblNum.Caption = Left(frmInvoice.lblNum.Caption, Len(Trim(frmInvoice.lblNum.Caption)) - 1)
            End If
        End If
    Else
        If LCase(VarActForm) = "frminvoice" Then
            frmInvoice.lblNum.Caption = frmInvoice.lblNum.Caption + BtnKeypad(index).Tag
        End If
    End If
End Sub

Private Sub BtnMenu_Click(index As Integer)
        
'    DeleteOptions
'    mvarcode = BtnMenu(Index).Tag
'    If mvarcode > 0 Then
'        If LCase(VarActForm) = "frminvoice" Then
'            Dim frmact, varForm As Form
'            For Each varForm In Forms
'                If LCase(VarActForm) = LCase(varForm.Name) Then
'                    Set frmact = varForm
'                    If frmact.GetGoodCode(mvarcode) = True Then
'                        frmact.ChangeGoodquantity
'                        frmact.lblNum.Caption = ""
'                    End If
'                End If
'            Next
'        End If
'    Else
'        mvarcode = 0
'    End If
'    If clsStation.MenuViewAfterGood = False Then
'        Unload Me
'    Else
'        FillGoodOptions BtnMenu(Index).Tag
'    End If
End Sub
Private Sub DeleteOptions()
    Dim ii As Long
    For ii = 0 To BtnOption.Count - 1
        BtnOption(ii).Caption = ""
        BtnOption(ii).Tag = 0
        BtnOption(ii).Enabled = False
        CostOption(ii) = 0
    Next
End Sub
Private Sub FillGoodOptions(GoodCode As Double)
    Dim Rst As New ADODB.Recordset
    Dim ii As Long
    ReDim Parameter(1) As Parameter
    
    Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, GoodCode)
    Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Difference", Parameter)
    ii = 0
    Do While Rst.EOF <> True
        If ii < BtnOption.Count Then    ''30 item
            BtnOption(ii).Caption = Rst!Difference
            BtnOption(ii).Tag = Rst!Code
            CostOption(ii) = Rst!CostDifference
            BtnOption(ii).Enabled = True
            ii = ii + 1
        Else
            Exit Do
        End If
        Rst.MoveNext
    Loop
    Set Rst = Nothing
End Sub

Private Sub BtnOption_Click(index As Integer)
    On Error GoTo ErrHandler
    mvarcode = Val(BtnOption(index).Tag)
    If mvarcode = 0 Then Exit Sub
    BtnOption(index).Enabled = False
    If clsStation.HasOptionPrice = True Then DisableCategoryOptions mvarcode
    If mvarcode > 0 Then
        If LCase(VarActForm) = "frminvoice" Then
            Dim frmact, varForm As Form
            For Each varForm In Forms
                If LCase(VarActForm) = LCase(varForm.Name) Then
                    Set frmact = varForm
                    frmact.FlxDetail.Select frmact.MaxRowFlexGrid - 1, 1
                    frmact.SplitByOptions frmact.FlxDetail.Row
                    If Len(frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 9)) <> 0 Then
                        frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 9) = frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 9) & ";"
                        frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 10) = frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 10) & ","
                    End If
                    frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 9) = frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 9) & BtnOption(index).Tag
                    frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 10) = frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 10) & BtnOption(index).Caption
                    If Right(frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 9), 1) = ";" Then
                        frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 9) = Left(frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 9), Len(frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 9)) - 1)
                        frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 10) = Left(frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 10), Len(frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 10)) - 1)
                    End If
                    frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 3) = frmact.FlxDetail.TextMatrix(frmact.FlxDetail.Row, 3) + CostOption(index)
                    frmact.RefreshLables
                    frmact.lblNum.Caption = ""
                    Exit For
                End If
            Next
        End If
    End If
    mvarcode = 0
Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 1500
End Sub
Private Sub DisableCategoryOptions(ByRef mvarcode)
    On Error GoTo ErrHandler
    Dim cnn As New ADODB.Connection
    Dim rctmp As New Recordset
    Dim strtemporary As String
    Dim ii As Long
    cnn.Open strConnectionString
    strtemporary = "SELECT * FROM dbo.tDifferences WHERE CategoryType = (SELECT CategoryType FROM dbo.tDifferences WHERE Code = " & mvarcode & ")"
    rctmp.Open strtemporary, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While Not rctmp.EOF
            For ii = 0 To BtnOption.Count - 1
                If rctmp!Code = BtnOption(ii).Tag Then
                    BtnOption(ii).Enabled = False
                    Exit For
                End If
            Next
            rctmp.MoveNext
        Loop
    End If
    rctmp.Close
    cnn.Close
    Set cnn = Nothing
Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 1500

End Sub

Private Sub cmdEscape_Click()
     Unload Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    'ElseIf KeyCode = 13 Then
        'OKButton_Click
    ElseIf KeyCode >= 47 And KeyCode <= 57 Then
        If LCase(VarActForm) = "frminvoice" Then
            frmInvoice.lblNum.Caption = frmInvoice.lblNum.Caption + Chr(KeyCode)
        End If
    ElseIf KeyCode = 8 Then
        If LCase(VarActForm) = "frminvoice" Then
            If Len(Trim(frmInvoice.lblNum.Caption)) >= 1 Then
                frmInvoice.lblNum.Caption = Left(frmInvoice.lblNum.Caption, Len(Trim(frmInvoice.lblNum.Caption)) - 1)
            End If
        End If
        
    
    End If
End Sub

Private Sub Form_Load()

   ' CenterCenterinSecondScreen Me
    If Screen.Width > 12000 Then
        Me.Left = 7700
        Me.Top = 1000
    Else
        Me.Left = 4000
        Me.Top = 1000
    End If
    
    If LCase(VarActForm) = "frminvoice" Then
        NotSupportedGoodType = EnumGoodType.forBuy
    End If
    
    mvarcode = 0
     
'    FillGoods (mvarGoodCode)
    formloadFlag = False
    Me.Left = Val(GetSetting(strMainKey, Me.Name, "Left"))
''''    If Val(GetSetting(strMainKey,Me.Name, "Height")) > 5000Then
''''        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
''''    End If
''''    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 5000 Then
''''        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
''''    End If
    Me.Top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    If Me.Left < 0 Then Me.Left = 0
    If Me.Top < 0 Then Me.Top = 0
    If Me.Top > Me.ScaleHeight Then Me.Top = 0

    formloadFlag = True
    For i = 0 To BtnOption.Count - 1
        Select Case clsStation.Language
            Case EnumLanguage.Farsi
                BtnOption(i).Font.Name = Invoice_FontMenuName
                BtnOption(i).Font.Size = Val(Invoice_FontMenuSize)
                BtnOption(i).Font.Bold = Invoice_FontMenuBold
            Case EnumLanguage.English
                BtnOption(i).Font = "TimesNewRoman"
                BtnOption(i).Font.Size = 10
                BtnOption(i).Font.Bold = True
        End Select
    Next i
    FillGoodOptions mvarGoodCode

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    SaveSetting strMainKey, Me.Name, "Height", Me.Height
    SaveSetting strMainKey, Me.Name, "Width", Me.Width
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


End Sub

Private Sub FillGoods(ByRef index)
'    Dim strStream
'    Dim rctmp As New ADODB.Recordset
'    ReDim Parameter(3) As Parameter
'
'    Parameter(0) = GenerateInputParameter("@BtnNum", adInteger, 4, mvarIndex)
'    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
'    Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
'    Parameter(3) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
'
'    Set rctmp = RunParametricStoredProcedure2Rec("GetGoodList", Parameter)
'    i = 0
'    Dim filetemp  As New FileSystemObject
'    Do While Not rctmp.EOF
'        BtnMenu(i).Tag = rctmp.Fields("GoodCode")
'        BtnMenu(i).Caption = rctmp.Fields("Name")
'        If IsNull(rctmp.Fields("Picture").Value) Then
'            If filetemp.FileExists(App.Path & Trim(rctmp.Fields("PicturePath"))) Then
'                BtnMenu(i).Picture = LoadPicture(App.Path & Trim(rctmp.Fields("PicturePath")))
'                'BtnMenu(i).WordWrap = False
'            Else
'                BtnMenu(i).Picture = LoadPicture("")
'            End If
'        Else
'            Set strStream = New ADODB.Stream
'            strStream.Type = adTypeBinary
'            strStream.Open
'            strStream.Write rctmp.Fields("Picture").Value
'            strStream.SaveToFile "C:\Temp.bmp", adSaveCreateOverWrite
'            BtnMenu(i).Picture = LoadPicture("C:\Temp.bmp")
'            Kill ("C:\Temp.bmp")
''            LoadPictureFromDB = True
'            Set strStream = Nothing
'        End If
'        i = i + 1
'        rctmp.MoveNext
'        If i = 35 Then Exit Do
'    Loop
'    rctmp.Cancel
'    i = i - 1
'    Dim h As Integer
''    h = 4600
'    h = 500
'    If i <= 6 Then
'        frameMenu.Height = h + 1000
'    ElseIf i >= 7 And i <= 13 Then
'        frameMenu.Height = h + 2000
'    ElseIf i >= 14 And i <= 20 Then
'        frameMenu.Height = h + 3000
'   ElseIf i >= 21 And i <= 27 Then
'        frameMenu.Height = h + 4000
'    ElseIf i >= 28 And i <= 34 Then
'        frameMenu.Height = h + 5000
''    ElseIf i >= 25 And i <= 29 Then
''        frameMenu.Height = h + 4500
''        Me.Height = h + 4600
''    ElseIf i >= 30 And i <= 34 Then
''        frameMenu.Height = h + 5400
''        Me.Height = h + 5500
''    ElseIf i >= 35 And i <= 39 Then
''        frameMenu.Height = h + 6300
''        Me.Height = h + 6400
'    End If
''    Me.Height = frameMenu.Height + 1400
'    For i = 1 To BtnMenu.Count - 1
'        If BtnMenu(i).Tag = "" And BtnMenu(i).Caption = "" Then
'            BtnMenu(i).Enabled = False
'         '   BtnMenu(i).WordWrap = True     ' Single Line
'            BtnMenu(i).Visible = False
'        End If
'    Next i
'
'    Frame_Option.Top = frameMenu.Height + 1300
'    Me.Height = frameMenu.Height + Frame_Option.Height + 1800
'    Set rctmp = Nothing
End Sub



