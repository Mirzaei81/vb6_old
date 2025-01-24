VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{75D4F148-8785-11D3-93AD-0000832EF44D}#4.0#0"; "FAST2003.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{75D4F6FF-8785-11D3-93AD-0000832EF44D}#1.3#0"; "FAST20051.ocx"
Object = "{75D4F4A8-8785-11D3-93AD-0000832EF44D}#3.3#0"; "FAST2007.ocx"
Object = "{75D4F666-8785-11D3-93AD-0000832EF44D}#3.3#0"; "FAST2006.ocx"
Begin VB.MDIForm mdifrm 
   BackColor       =   &H00FF8080&
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   16095
   Icon            =   "mdifrm.frx":0000
   LinkTopic       =   "MDIForm1"
   RightToLeft     =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList7 
      Left            =   2400
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   36
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":B6A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":12D81
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1A45B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1D6B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":227D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":29EB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":3158B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":38C65
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":4033F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":47A19
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":4F0F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":567CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":59A27
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":5CC81
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":6435B
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":6BA35
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":6EC8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":71EE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":795C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":7C81D
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":83EF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":8B5D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":92CAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":9A385
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":A1A5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":A9139
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":B0813
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":B7EED
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":BF5C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":C6CA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":CE37B
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":D5A55
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":D8CAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":DBF09
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":DF163
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":E23BD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrScreenSaver 
      Interval        =   60000
      Left            =   0
      Top             =   960
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   9885
      Visible         =   0   'False
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin FLWSystem.FWSysTray objSysTray 
      Left            =   3240
      Top             =   960
      _ExtentX        =   926
      _ExtentY        =   926
      Icon            =   0
      Text            =   ""
   End
   Begin VB.Timer tmrUdp 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   480
      Top             =   1560
   End
   Begin FLWSystem.FWRegistry FWRegistry1 
      Left            =   2520
      Top             =   840
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin FLWDiskFile.FWDisks FWDisks1 
      Left            =   1800
      Top             =   840
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin FLWData.FWEncryption FWEncryption1 
      Left            =   1080
      Top             =   840
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1800
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   6345
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   64
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":E8C1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":E8F3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":E9257
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":E9571
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":E988B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":E9BA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":EA47F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":EAD59
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":EB633
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":EBF0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":EC7E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":ED0C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":ED99B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":EE275
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":EEB4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":EF429
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":EFD03
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":F05DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":F0EB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":F1791
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":F206B
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":F2945
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":F321F
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":F3AF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":F43D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":F4CAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":FB50F
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":101D71
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1085D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":10EE35
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":115697
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":11BEF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":12275B
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":128FBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":12F81F
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":136081
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":13C8E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":143145
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1499A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":150209
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":156A6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":15D2CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":163B2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":16A391
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":170BF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":177455
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":17DCB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":184519
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":18AD7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1915DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":197E3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":19E6A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1A4F03
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1AB765
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1B1FC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1B8829
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1B8EE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1B95E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1B9CE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1BA5BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1EA60D
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1F0E6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1F76D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":1FDF33
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6360
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   50
      ImageHeight     =   50
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":204795
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":204F39
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":2056DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":205E81
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":206625
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":206DC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20756D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":207D11
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":2084B5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5055
      Top             =   1770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5040
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":208C59
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":208EB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":209025
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":209191
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":2093F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":209689
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":209951
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":209BF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":209E9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20A301
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20A569
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20A7D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20C83D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20CAC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20CDF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20D1A9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   10515
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   16008
            MinWidth        =   14288
            Text            =   "    "
            TextSave        =   "    "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1826
            MinWidth        =   1411
            TextSave        =   "04:06 ».Ÿ"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2347
            MinWidth        =   2347
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Text            =   "”„ "
            TextSave        =   "”„ "
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Text            =   "‰«„ ﬂ«—»— "
            TextSave        =   "‰«„ ﬂ«—»— "
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "Ver : "
            TextSave        =   "Ver : "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   6360
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20D529
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20D845
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20DB61
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20DE7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20E199
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20E4B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20E7D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20EAED
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20EE09
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20F125
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20F441
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":20F75D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":210039
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":210915
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":210C31
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":210F4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":211269
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":211B45
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":211E61
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":214615
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":214939
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":214C5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":214F79
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":215293
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":2155AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":2158C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":215BE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":215EFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":216215
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":21652F
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":216E09
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":2176E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":217FBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":218897
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":219171
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":219A4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":21A325
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":21ABFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":21B4D9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock WinsockUdp 
      Left            =   2400
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin FLWMMedia.FWMMedia FWMMedia1 
      Left            =   0
      Top             =   1440
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin MSWinsockLib.Winsock Winsock_Print 
      Left            =   2400
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1560
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":21BDB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":21DA8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":21F767
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":221441
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":22311B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":224DF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":228997
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":22C539
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":2300DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":233C7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":23781F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":23B3C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":23EF63
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":242B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":2466A7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList6 
      Left            =   3240
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":248381
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":24A05B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":24BD35
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":24DA0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":24F6E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":2513C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":254F65
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":258B07
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":25C6A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":26024B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":263DED
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":26798F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":26B531
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":26F0D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":272C75
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdifrm.frx":27494F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   1905
      ButtonWidth     =   1561
      ButtonHeight    =   1852
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList6"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   30
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«» œ«"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Key             =   "dgvsd"
                  Object.Tag             =   "1"
                  Text            =   "text"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ﬁ»·Ì"
            Object.ToolTipText     =   "Page Down"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "»⁄œÌ"
            Object.ToolTipText     =   "Page Up"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«‰ Â«"
            Object.ToolTipText     =   "End"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«›“Êœ‰"
            Object.ToolTipText     =   "Ins"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ÊÌ—«Ì‘"
            Object.ToolTipText     =   "F3"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "À» "
            Object.ToolTipText     =   "Enter"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«‰’—«›"
            Object.ToolTipText     =   "Esc"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Õ–›"
            Object.ToolTipText     =   "Del"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "—«Â‰„«"
            Object.ToolTipText     =   "F1"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ã” ÃÊ"
            Object.ToolTipText     =   "F2"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "À»  Ê ç«Å"
            Object.ToolTipText     =   "F6"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "«”ò‰"
            Object.ToolTipText     =   "F12"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "„—ÃÊ⁄"
            Object.ToolTipText     =   "„—ÃÊ⁄"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   4
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "›Ê‰ "
            Object.ToolTipText     =   "«‰ Œ«» ›Ê‰ "
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "    —‰ê   "
            Object.ToolTipText     =   "«‰ Œ«» —‰ê "
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " “»«‰"
            Key             =   "English"
            Object.ToolTipText     =   "«‰ Œ«» “»«‰"
            ImageIndex      =   16
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "  ·›‰"
            Object.ToolTipText     =   "Ctrl+F9"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "òÌ»Ê—œ"
            Key             =   "KeyBoard"
            Object.ToolTipText     =   "òÌ »—œ"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "„ Õ”«»"
            Key             =   "Calculator"
            Object.ToolTipText     =   "„«‘Ì‰ Õ”«»"
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " —ò ’"
            Object.ToolTipText     =   "F11"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Œ—ÊÃ"
            Object.ToolTipText     =   "Ctrl+F12"
            ImageIndex      =   4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dsgfs"
                  Object.Tag             =   "sdgsdfg"
                  Text            =   "text"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
   End
   Begin MSWinsockLib.Winsock Winsock_Farabin 
      Left            =   480
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Total.TelnetTTYClient ttcControl 
      Left            =   480
      Top             =   3240
      _ExtentX        =   873
      _ExtentY        =   873
   End
End
Attribute VB_Name = "mdifrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Msg As String
Private s As String
Public ClsActionLog As New ClsActionLog
Private clsDate As New clsDate
Private rctmp As New ADODB.Recordset
Private cmd As New ADODB.command
Dim Parameter() As Parameter
Dim filetemp As New FileSystemObject
Private WithEvents TinyEvent As TINYLib.Tiny
Attribute TinyEvent.VB_VarHelpID = -1
Dim LineNumberTemp As Long
Private Const APPCAPTION As String = "TelnetTTY"
Private Const LFCR = vbLf & vbCr

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub
''''Private Sub fwBtnCtrl_Click()
''''    Call PresetScreenSaver
''''' Call NeccesaryFunction
''''    Dim Result As Boolean
''''    Dim frmActive As Form
'''''    If clsStation.TreeViewMenu Then
'''''        Result = False
'''''        For Each frmActive In Forms
'''''            If frmActive.Name = "frmGroupMenu" Then
'''''                Result = True
'''''                Exit For
'''''            End If
'''''        Next
'''''        If Result = False Then frmGroupMenu.Show
'''''        Me.WindowState = vbMaximized
'''''    Else
''''        If Toolbar3.Visible = False Then
''''            Toolbar3.Visible = True
''''            If LCase(VarActForm) = "frminvoice" Or LCase(VarActForm) = "frmpurchase" Then
''''                mdifrm.Toolbar3.Buttons(1).Enabled = False
''''                mdifrm.Toolbar3.Buttons(2).Enabled = False
''''            End If
''''        Else
''''            Toolbar3.Visible = False
''''        End If
'''''    End If
''''End Sub

''''Private Sub fwBtnCtrl_DblClick()
''''    fwBtnCtrl_Click
''''End Sub
''''Private Sub fwBtnCtrl_KeyUp(KeyCode As Integer, Shift As Integer)
''''Dim test As Double
''''
''''Select Case Shift
''''    Case 0:
''''        Select Case KeyCode
''''            Case 27: 'Esc
''''                If clsStation.KeyboardType = EnumKeyBoardType.S1 Then
''''                    ShowMessage "¬Ì« »—«Ì Œ—ÊÃ «“ »—‰«„Â «ÿ„Ì‰«‰ œ«—Ìœ", True, True, "»·Ì", "ŒÌ—"
''''                    If mvarMsgIdx = vbYes Then
''''                        Unload Me
''''                    End If
''''                End If
''''
''''            Case 122:   'Screen Saver
''''                'test = Shell(App.Path & "\Tools\OSA.EXE  -s", vbMaximizedFocus)
''''                'If test = 0 Then
''''                '    GoTo Err1
''''                'End If
''''                Unload Me
'''''            Case 121:
'''''                Unload Me
''''        End Select
''''    Case 1:
''''
''''    Case 2: 'Control Key Press
''''        Select Case KeyCode
''''            Case 119:       ' Cash Drawer Open
''''
''''            Case 123:    '(Ctl + f12)
''''            If clsStation.KeyboardType = EnumKeyBoardType.Rb2 Or clsStation.KeyboardType = EnumKeyBoardType.Promag Then
''''                ShowMessage "¬Ì« »—«Ì Œ—ÊÃ «“ »—‰«„Â «ÿ„Ì‰«‰ œ«—Ìœ", True, True, "»·Ì", "ŒÌ—"
''''                If mvarMsgIdx = vbYes Then
''''                    Unload Me
''''                End If
''''            End If
''''        End Select
''''End Select
''''If KeyCode <> 123 Or KeyCode <> 27 Then
''''    mdifrm.fwBtnCtrl.SetFocus
''''End If
''''
''''Exit Sub
''''
''''Err1:
''''    ShowMessage " Screen Saver ”Ì” „ Ì«›  ‰‘œ", True, False, " «ÌÌœ", " "
''''    Unload Me
''''    FrmLogin.Show
''''End Sub
'''''


Private Sub MDIForm_Activate()
    
    If clsArya.HardLockSerialNo = "93061701000" Then
        mdifrm.Caption = Space(100) & Trim(clsArya.Company) & "  -  " & clsArya.DbName
    Else
        mdifrm.Caption = Space(120) & Trim(clsArya.Company)
    End If
    If VarActForm = "" Then
        Toolbar1.Buttons(23).Enabled = False
        Toolbar1.Buttons(24).Enabled = True
        Toolbar1.Buttons(25).Enabled = False
    '    Toolbar1.Buttons(27).Enabled = False
    Else
        Toolbar1.Buttons(23).Enabled = True
        Toolbar1.Buttons(24).Enabled = True
        Toolbar1.Buttons(25).Enabled = True
        Toolbar1.Buttons(27).Enabled = True
    End If
    If clsStation.TextIconViewH = False Then
            Dim i As Integer
            For i = 1 To 29
                Toolbar1.Buttons(i).Caption = ""
            Next i
    End If
    
    Unload frmfactor
End Sub

Private Sub MDIForm_Deactivate()
        'Me.mnuBasInfo.Visible = False
End Sub

Private Sub MDIForm_Resize()
    frmGroupMenu.Height = Me.Height - 300
    frmAbout.Left = (Me.Width - frmGroupMenu.Width - frmAbout.Width) / 2
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error GoTo Err_Handler
    If clsStation.TelNetServerActive = True Then
        ttcControl.Disconnect
        Sleep 500
        Unload frmTerminal
    End If
   '========= Auto BackUp ===========
    If clsStation.AutoBackup = True Then
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.lblMessage = "”Ì” „ œ— Õ«· ê—› ‰ ‰”ŒÂ Å‘ Ì»«‰ «“ œ«œÂ Â« „Ì »«‘œ "
        frmDisMsg.lblMessage = frmDisMsg.lblMessage & vbLf & "·ÿ›« „‰ Ÿ— »„«‰Ìœ "
        frmDisMsg.Show
        DoEvents
        On Error Resume Next
        AutoBackup
        On Error GoTo 0
   
    End If
    If clsArya.PrintServer = True And Station_IsServer = True Then
        mdifrm.Winsock_Print.SendData "Exit"
        CloseWindow "PrintServer"
    End If
    VarActForm = ""
    CloseWindow "On-Screen Keyboard"
    Unload frmGroupMenu
    Set clsArya = Nothing
    Set clsStation = Nothing
    SetKbLayout LANG_EN_US
    
    Call objSysTray.Destroy
    
   'Insert user logout history
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@intUserNo", adInteger, 4, mvarCurUserNo)
    Parameter(1) = GenerateInputParameter("@intActionUserNo", adInteger, 4, 2) '2 means log out
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    RunParametricStoredProcedure "Insert_tblTotal_UserHistory", Parameter
                
    If PosConnection.State = 1 Then PosConnection.Close
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    If HardLockFlag = True Or HardLockFlagTrial = True Then
        Tiny1.UserPassWord (KarbarKey)
        Tiny1.SetAutoCheckingTinyHID (False)
        Tiny1.DisconnectFromTinyHID
    End If
    
    If SecurityVersion = 1 Then
        strSockRecive = ""
        If Winsock1.State = sckConnected Then
            Winsock1.SendData Operations.LogOutStation & seperator & EOS
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
    SetKbLayout LANG_EN_US
    End

Exit Sub
Err_Handler:
    LogSaveNew "MDIForm => ", err.Description, err.Number, err.Source, "MDIFrm_Unload"
    ShowErrorMessage
End Sub
Public Sub ExitMdiForm()
    
      frmMsg.fwlblMsg.Caption = "¬Ì« »—«Ì Œ—ÊÃ «“ »—‰«„Â «ÿ„Ì‰«‰ œ«—Ìœ"
    ' frmMsg.fwBtn(0).Visible = True
      frmMsg.fwBtn(0).ButtonType = flwButtonOk
      frmMsg.fwBtn(1).ButtonType = flwButtonCancel
      frmMsg.fwBtn(0).Caption = "»·Ì"
      frmMsg.fwBtn(1).Caption = "ŒÌ—"
      frmMsg.Show vbModal
      If mvarMsgIdx = vbYes Then
        Unload Me
       ' End
      End If

End Sub

Private Sub MDIForm_Load()
    On Error GoTo Err_Handler
    
    Dim hMenu As Long

    hMenu = GetSystemMenu(Me.hWnd, False)

    DeleteMenu hMenu, 6, MF_BYPOSITION

    AllButton vbOff, True
    
'    Skin1.LoadSkin App.Path & "\Skins\metallic.skn"
'    Skin1.ApplySkin Me.Hwnd
        
    Toolbar1.Buttons(20).Enabled = False
    Toolbar1.Buttons(21).Enabled = False
    Toolbar1.Buttons(23).Enabled = True
    Toolbar1.Buttons(24).Enabled = True
    Toolbar1.Buttons(25).Enabled = True
    Toolbar1.Buttons(27).Enabled = True
    
   
'    If clsArya.MaxStationNo = 1 Then SecurityVersion = 0 'Bridge Not Need to Run
    
    If SecurityVersion = 1 Then
        modsock.ConnectSock
    End If
    
    If clsStation.NetworkCallerId = True Then
'       tmrUdp.Enabled = True
       modsock.ConnectToClient
    End If
    
    If clsStation.TelNetServerActive = True Then
        ttcControl.TermType = "NVT|TTY"
        ttcControl.Connect clsStation.TelNetServerIP, clsStation.TelNetServerPort
        LineNumberTemp = 0
       ' frmTerminal.txtInput.SetFocus
    End If
    
    If clsStation.Language = English Then
        SetKbLayout LANG_EN_US
        Toolbar1.Buttons(23).Value = tbrPressed
    Else
        If DebugMode = False Then SetKbLayout LANG_Pr_IR
    End If
    
'    mdifrm.fwBtnCtrl.Left = Screen.Width - 650 ' 14640

    RepVer = "_V26_16"
     
'    If filetemp.FileExists(App.Path & "\Tools\Clock.exe") And DebugMode = False Then
'        Shell App.Path & "\Tools\Clock.exe", vbNormalNoFocus
'    End If
      objSysTray.Icon = Me.Icon
'      objSysTray.About
      objSysTray.Text = "”Ì” „ —” Ê—«‰ ”„— :  Ê”ÿ «› ÃÌ ¬—Ì« "
      
    Call objSysTray.Create
    
    Call NeccesaryFunction

    ChangeToolBar1Language
    ChangeToolbar3Language
    Call clsArya.VersionDefine

    If clsStation.PosPayment = True And clsStation.PosModel > 0 Then
''        PosSocketInit
    End If
    
    If clsArya.PrintServer = True And Station_IsServer = True Then
        CloseWindow "PrintServer"
        If filetemp.FileExists(App.Path & "\Printing\Print_Server.exe") Then
            Shell App.Path & "\Printing\Print_Server_V26_16.exe", vbMinimizedFocus
            Sleep 100
            modsock.ConnectToClient2
        Else
            MsgBox " ›«Ì· ÅÌœ« ‰‘œ " & App.Path & "\Printing\Print_Server_V26_16.exe"
            clsArya.PrintServer = False
        End If
    End If
    
    If IsFarabin = True Then
        With mdifrm.Winsock_Farabin
            .Close
            .RemoteHost = .LocalIP
            .RemotePort = 5200
        End With
        If filetemp.FileExists(App.Path & "\DigitalPlayerMini.exe") Then
            Shell App.Path & "\DigitalPlayerMini.exe", vbMinimizedFocus
            'Shell App.Path & "\Server.exe", vbMinimizedFocus
           ''' modsock.ConnectToClient_Farabin  No Need Connect For initialize Because connecting from into invoice
        Else
            ShowDisMessage " ›«Ì· ÅÌœ« ‰‘œ " & App.Path & "\DigitalPlayerMini.exe", 1500
        End If
    End If
    
    '' because will be checked before GroupMenu
    If (clsArya.MaxAccountingNo > 0 And Station_IsAccounting = True) Then       ''
        CheckAccounting
    ElseIf HasMiniAcc = True Then        ''
        CheckAccounting
        clsArya.ExternalAccounting = False
    Else
        clsArya.ExternalAccounting = False
    End If
    
    If HasPcPos = True Then CheckPcPos
    
    frmGroupMenu.Height = Me.Height - 300
    If HardLockFlag = True Or HardLockFlagTrial = True Then
        Set TinyEvent = Tiny1
    End If
'
    Exit Sub
    
Err_Handler:
    If err.Number = 87 Then
        MsgBox "¬œ—” ¬Ì ÅÌ  ‰ŸÌ„ ‘œÂ œ—  ‰ŸÌ„«  »—«Ì »—ﬁ—«—Ì «— »«ÿ »« ﬂ«·— ¬Ì œÌ ‰ê«— «‘ »«Â «” . ·ÿ›« ¬‰ —« «ÿ·«Õ Ê »—‰«„Â —« œÊ»«—Â «Ã—« ‰„«ÌÌœ"
        err.Clear
    Else
        ShowErrorMessage
        LogSaveNew "mdifrm => ", err.Description, err.Number, err.Source, "MDIForm_Load"
        Resume Next
    End If

End Sub

Private Sub ttcControl_Connect()
    ShowDisMessage "« ’«· »Â ”—Ê—  · ‰  «‰Ã«„ ‘œ", 1000
   ' ttcControl.Echo False
    frmTerminal.Show vbModeless, Me
  '  frmTerminal.txtInput.SetFocus
End Sub

Private Sub ttcControl_DataArrival()
    
    On Error GoTo ErrHandler
    Dim strData As String
    
    strData = Replace$(ttcControl.GetData(), vbCrLf, vbFormFeed)
    strData = Replace$(strData, LFCR, vbFormFeed)
    strData = Replace$(strData, vbLf, vbFormFeed)
    strData = Replace$(strData, vbFormFeed, vbCrLf)
    With frmTerminal.txtLog
        If Len(.Text) > 10000 Then
            .Text = Right$(.Text, 10000)
        End If
        .Text = .Text & strData
        .SelStart = Len(.Text)
    End With
     
    '=======Added for caller id output
    'Output line fot Telnet is below (in strData string)
    '**CID                |0  |1  |16  |9123093493|fd892ab7-66be-43a6-a856-0d50466c15bd.
    ' we must get caller id from this line
    Dim CidParameter As Long
    Dim Column1, Column2, Column3, column4, column5, LineNumber As Long
    
    CidParameter = InStr(1, strData, "**CID", vbBinaryCompare)
    If CidParameter > 0 Then
        Column1 = InStr(CidParameter + 1, strData, "|", vbBinaryCompare)
    End If
    If Column1 > 0 Then
        Column2 = InStr(Column1 + 1, strData, "|", vbBinaryCompare)
    End If
    If Column2 > 0 Then
        Column3 = InStr(Column2 + 1, strData, "|", vbBinaryCompare)
    End If
    If Column3 > 0 Then
        column4 = InStr(Column3 + 1, strData, "|", vbBinaryCompare)
    End If
    If column4 > 0 Then
        LineNumber = Val(Mid(strData, Column3 + 1, column4 - Column3))
        column5 = InStr(column4 + 1, strData, "|", vbBinaryCompare)
    End If
    Dim Call_NumberTemp As String
    If column5 > 0 Then
        With frmTerminal.txtCallerId
            If Len(.Text) > 10000 Then
                .Text = Right$(.Text, 1000)
            End If
            Call_NumberTemp = Mid(strData, column4 + 1, column5 - column4)
            .Text = .Text & Call_NumberTemp + vbCrLf
            .SelStart = Len(.Text)
        End With
        If (Len(Call_NumberTemp) > 8 And Left(Call_NumberTemp, 1) <> "0") Then Call_NumberTemp = "0" & Call_NumberTemp   ' call from other city
'        WinsockUdp.SendData "101" & Call_NumberTemp
        
        ' because has 40 line and we must set it to max 8 line
        LineNumberTemp = LineNumberTemp + 1
        If LineNumberTemp = 9 Then LineNumberTemp = 8
    'Save CallerId In Database
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@nvcDate", adWChar, 8, mvarDate)
        Parameter(1) = GenerateInputParameter("@LineNumber", adTinyInt, 1, LineNumber)
        Parameter(2) = GenerateInputParameter("@nvcCallerId", adWChar, 20, Call_NumberTemp)
        RunParametricStoredProcedure "Insert_tblTotal_CallerId", Parameter

        If clsStation.NetworkCallerId = True Then
            mdifrm.WinsockUdp.SendData Str(LineNumber) & "0" & Call_NumberTemp
          '  LogSaveNew Inputstr, "", "", "", ""
            LogSaveNew "Network Send: " & Str(LineNumber) & "0" & Call_NumberTemp, "", "", ""
'                Sleep 200
'                mdifrm.WinsockUdp.SendData Str(LineNumber) & Str(Index) & Call_NumberTemp
        End If
        
    End If
    '================================
Exit Sub
ErrHandler:
    MsgBox err.Description
    
End Sub

Private Sub ttcControl_Disconnect()
    ShowDisMessage "« ’«· ”—Ê—  ·‰  ﬁÿ⁄ ‘œ .", 1000
    'ttcControl.Disconnect
End Sub

Private Sub ttcControl_Error(ByVal Number As Long, ByVal Description As String)
'    MsgBox "Error &H" & Hex$(Number) & " " & Description, _
'           vbOKOnly Or vbExclamation, APPCAPTION
    ShowDisMessage "Œÿ« œ— ”—Ê—  ·‰ " & err.Description, 1000
End Sub

Private Sub CheckPcPos()
    
    Dim tempstring As TextStream
    If filetemp.FileExists(App.Path & "\AryaPosData.dll") = True Then
        If filetemp.FileExists(App.Path & "\AryaPosData.tlb") = False Then
            If FWRegistry1.KeyExists(flwRegLocalMachine, "SOFTWARE\Microsoft\.NETFramework") Then
                Dim InstallPath As String
                Call FWRegistry1.GetKeyStr(flwRegLocalMachine, "SOFTWARE\Microsoft\.NETFramework", "InstallRoot", InstallPath)

                If filetemp.FileExists(App.Path & "\Register_tlb.cmd") = True Then
                    filetemp.DeleteFile App.Path & "\Register_tlb.cmd", True
                End If

                filetemp.CreateTextFile App.Path & "\Register_tlb.cmd"
                Set tempstring = filetemp.OpenTextFile(App.Path & "\Register_tlb.cmd", ForWriting, False, TristateFalse)
                If filetemp.FolderExists(InstallPath & "v4.0.30319") = True Then
                    tempstring.WriteLine "C:"
                    tempstring.WriteLine "cd " & InstallPath & "v4.0.30319"
                    tempstring.WriteLine "RegAsm.exe " & App.Path & "\AryaPosData.dll /CodeBase /tlb"
                    tempstring.Close
                    Shell App.Path & "\Register_tlb.cmd", vbNormalFocus
                ElseIf filetemp.FolderExists(InstallPath & "v2.0.50727") = True Then
                    tempstring.WriteLine "C:"
                    tempstring.WriteLine "cd " & InstallPath & "v2.0.50727"
                    tempstring.WriteLine "RegAsm.exe " & App.Path & "\AryaPosData.dll /CodeBase /tlb"
                    tempstring.Close
                    Shell App.Path & "\Register_tlb.cmd", vbNormalFocus
                Else
                    MsgBox ("ÅÊ‘Â" & ".NetFramework" & "»—«Ì «” ›«œÂ œ— ÅÊ“ »«‰òÌ ÅÌœ« ‰‘œ" & vbLf & "”Ì” „ ›—Ê‘ ﬁ«œ— »Â »—ﬁ—«—Ì «— »«ÿ ¬‰·«Ì‰ »« ÅÊ“ »«‰òÌ ‰Ì” " & vbLf & "›«Ì· „Ê—œ ‰Ÿ— —« «“ œ«Œ· ”Ì œÌ ‰’» «Ã—« Ê ‰’» ﬂ‰Ìœ")                 'or permission code goes here
                    HasPcPos = False
                End If
            Else
                MsgBox (".NetFramework" & "»—«Ì «” ›«œÂ œ— ÅÊ“ »«‰òÌ ÅÌœ« ‰‘œ" & vbLf & "”Ì” „ ›—Ê‘ ﬁ«œ— »Â »—ﬁ—«—Ì «— »«ÿ ¬‰·«Ì‰ »« ÅÊ“ »«‰òÌ ‰Ì” ")               'or permission code goes here
                HasPcPos = False
            End If
        End If
    End If
End Sub
Private Sub CheckAccounting()
    
    Dim tempstring As TextStream
    clsArya.ExternalAccounting = True

    '''Register Dll
    
    If filetemp.FileExists(App.Path & "\prjAccount.dll") = True Then
'            If filetemp.FileExists(SystemFolderName & "\prjAccount.dll") = True Then
'                filetemp.DeleteFile SystemFolderName & "\prjAccount.dll", True
'            End If
'            filetemp.MoveFile App.Path & "\prjAccount.dll", SystemFolderName & "\prjAccount.dll"
        If filetemp.FileExists(App.Path & "\RegisterAc.cmd") = True Then
            filetemp.DeleteFile App.Path & "\RegisterAc.cmd", True
        End If
        filetemp.CreateTextFile App.Path & "\RegisterAc.cmd"
        Set tempstring = filetemp.OpenTextFile(App.Path & "\RegisterAc.cmd", ForWriting, False, TristateFalse)
        If filetemp.FileExists(SystemFolderName & "\prjAccount.dll") = True Then
            tempstring.WriteLine "regsvr32 " & SystemFolderName & "\prjAccount.dll /U"
            tempstring.WriteLine "Del " & SystemFolderName & "\prjAccount.dll "
        Else
            tempstring.WriteLine "regsvr32 " & App.Path & "\prjAccount.dll /U"
        End If
        tempstring.WriteLine "Move " & "prjAccount.dll " & SystemFolderName & "\prjAccount.dll "
        tempstring.WriteLine "regsvr32 " & SystemFolderName & "\prjAccount.dll "
        tempstring.Close
        Shell App.Path & "\RegisterAc.cmd", vbNormalFocus
        MsgBox "›«Ì· Õ”«»œ«—Ì ›⁄«· ‘œ"
    Else
        If filetemp.FileExists(SystemFolderName & "\prjAccount.dll") = False Then
            MsgBox (" ›«Ì·  prjAccount.dll  " & "»—«Ì «” ›«œÂ œ— Õ”«»œ«—Ì ÅÌœ« ‰‘œ" & vbLf & "”Ì” „ ›—Ê‘ ﬁ«œ— »Â »—ﬁ—«—Ì «— »«ÿ ¬‰·«Ì‰ »« Õ”«»œ«—Ì ‰Ì” " & vbLf & "”Ì” „ Õ”«»œ«—Ì €Ì— ›⁄«· „Ì ‘Êœ")                'or permission code goes here
            clsArya.ExternalAccounting = False
        End If
    End If
    '''
    If filetemp.FileExists(App.Path & "\crviewer.dll") = True And clsArya.ExternalAccounting = True Then
        If filetemp.FileExists(App.Path & "\RegisterCrViewer.cmd") = True Then
            filetemp.DeleteFile App.Path & "\RegisterCrViewer.cmd", True
        End If
        filetemp.CreateTextFile App.Path & "\RegisterCrViewer.cmd"
        Set tempstring = filetemp.OpenTextFile(App.Path & "\RegisterCrViewer.cmd", ForWriting, False, TristateFalse)
        If filetemp.FileExists(SystemFolderName & "\crviewer.dll") = True Then
            tempstring.WriteLine "regsvr32 " & SystemFolderName & "\crviewer.dll /U"
            tempstring.WriteLine "Del " & SystemFolderName & "\crviewer.dll "
        Else
            tempstring.WriteLine "regsvr32 " & App.Path & "\crviewer.dll /U"
        End If
        tempstring.WriteLine "Move " & "crviewer.dll " & SystemFolderName & "\crviewer.dll "
        tempstring.WriteLine "regsvr32 " & SystemFolderName & "\crviewer.dll "
        tempstring.Close
        Shell App.Path & "\RegisterCrViewer.cmd", vbNormalFocus
        MsgBox "›«Ì· «Ê· ê“«—‘«  Õ”«»œ«—Ì ›⁄«· ‘œ"
    Else
        If filetemp.FileExists(SystemFolderName & "\crviewer.dll") = False Then
            MsgBox (" ›«Ì·  crviewer.dll  " & "»—«Ì «” ›«œÂ œ— Õ”«»œ«—Ì ÅÌœ« ‰‘œ" & vbLf & "ê“«—‘«  œ— Õ”«»œ«—Ì €Ì— ›⁄«· «”  ")                'or permission code goes here
            clsArya.ExternalAccounting = False
        End If
    End If
    If filetemp.FileExists(App.Path & "\craxddrt.dll") = True And clsArya.ExternalAccounting = True Then
        If filetemp.FileExists(App.Path & "\Registercraxddrt.cmd") = True Then
            filetemp.DeleteFile App.Path & "\Registercraxddrt.cmd", True
        End If
        filetemp.CreateTextFile App.Path & "\Registercraxddrt.cmd"
        Set tempstring = filetemp.OpenTextFile(App.Path & "\Registercraxddrt.cmd", ForWriting, False, TristateFalse)
        If filetemp.FileExists(SystemFolderName & "\craxddrt.dll") = True Then
            tempstring.WriteLine "regsvr32 " & SystemFolderName & "\craxddrt.dll /U"
            tempstring.WriteLine "Del " & SystemFolderName & "\craxddrt.dll "
        Else
            tempstring.WriteLine "regsvr32 " & App.Path & "\craxddrt.dll /U"
        End If
        tempstring.WriteLine "Move " & "craxddrt.dll " & SystemFolderName & "\craxddrt.dll "
        tempstring.WriteLine "regsvr32 " & SystemFolderName & "\craxddrt.dll "
        tempstring.Close
        Shell App.Path & "\Registercraxddrt.cmd", vbNormalFocus
        MsgBox "›«Ì· œÊ„ ê“«—‘«  Õ”«»œ«—Ì ›⁄«· ‘œ"
    Else
        If filetemp.FileExists(SystemFolderName & "\craxddrt.dll") = False Then
            MsgBox (" ›«Ì·  craxddrt.dll  " & "»—«Ì «” ›«œÂ œ— Õ”«»œ«—Ì ÅÌœ« ‰‘œ" & vbLf & "ê“«—‘«  œ— Õ”«»œ«—Ì €Ì— ›⁄«· «”  ")                'or permission code goes here
            clsArya.ExternalAccounting = False
        End If
    End If

'    If DebugMode = True Then
'        Set Accounting = New prjAccount.ClsMonitoring
'    Else
        Set Accounting = CreateObject("prjAccount.ClsMonitoring")
       ' Set Accounting = New prjAccount.ClsMonitoring
         ''Please Set modgl routine
'    End If
    
End Sub


Private Sub objSysTray_DblClk(Button As Integer)
    If Button = vbLeftButton Then
      ' show the window
      If Me.WindowState = vbNormal Then
          Me.WindowState = vbMaximized
      Else: Me.WindowState = vbNormal
      End If
      If Not Me.Visible Then
        Me.Show
      End If
    End If
    frmBrowser.Show
    Dim st As String
    If strDelegate = "56" Then st = "http://www.MoeinReklam.com" Else st = "http://www.fgarya.com/pages/43"
    frmBrowser.cboAddress.Text = st
    frmBrowser.cboAddress_Click
End Sub

Private Sub objSysTray_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
      ' show the menu
    '  Me.PopupMenu frmAbout.Show
      frmAbout.Show
    End If
    frmBrowser.Show
    Dim st As String
    If strDelegate = "56" Then st = "http://www.MoeinReklam.com" Else st = "http://www.fgarya.com/pages/43"
    frmBrowser.cboAddress.Text = st
    frmBrowser.cboAddress_Click
End Sub

Private Sub ChangeToolbar3Language()
'If clsStation.Language = English Then
'    Toolbar3.Buttons(1).Caption = "Sale"
'    Toolbar3.Buttons(2).Caption = "Buy"
'    Toolbar3.Buttons(3).Caption = "Accounting"
'    Toolbar3.Buttons(4).Caption = "Receive"
'    Toolbar3.Buttons(5).Caption = "Delivery"
'    Toolbar3.Buttons(6).Caption = ""
'    Toolbar3.Buttons(7).Caption = "Persons"
'    Toolbar3.Buttons(8).Caption = "Definitions"
'    Toolbar3.Buttons(9).Caption = "Setting"
'    Toolbar3.Buttons(10).Caption = "Facilities"
'    Toolbar3.Buttons(11).Caption = "Report"
'    Toolbar3.Buttons(12).Caption = ""
'    Toolbar3.Buttons(13).Caption = "View"
'    Toolbar3.Buttons(14).Caption = ""
'    Toolbar3.Buttons(15).Caption = "Access"
'    Toolbar3.Buttons(16).Caption = "Goods"
'    Toolbar3.Buttons(17).Caption = "Escape"
'    Toolbar3.Buttons(18).Caption = ""
'Else
'    Toolbar3.Buttons(1).Caption = "›—Ê‘"
'    Toolbar3.Buttons(2).Caption = "Œ—Ìœ"
'    Toolbar3.Buttons(3).Caption = "Õ”«»œ«—Ì"
'    Toolbar3.Buttons(4).Caption = "œ—Ì«› "
'    Toolbar3.Buttons(5).Caption = "ÅÌò"
'    Toolbar3.Buttons(6).Caption = ""
'    Toolbar3.Buttons(7).Caption = "«‘Œ«’"
'    Toolbar3.Buttons(8).Caption = " ⁄«—Ì›"
'    Toolbar3.Buttons(9).Caption = " ‰ŸÌ„« "
'    Toolbar3.Buttons(10).Caption = "«„ò«‰« "
'    Toolbar3.Buttons(11).Caption = "ê“«—‘« "
'    Toolbar3.Buttons(12).Caption = ""
'    Toolbar3.Buttons(13).Caption = "‰„«Ì‘"
'    Toolbar3.Buttons(14).Caption = ""
'    Toolbar3.Buttons(15).Caption = "œ” —”Ì"
'    Toolbar3.Buttons(16).Caption = "«‰»«— Ê ò«·«"
'    Toolbar3.Buttons(17).Caption = " —ò ’‰œÊﬁ"
'    Toolbar3.Buttons(18).Caption = ""
'End If


End Sub
Private Sub ChangeToolBar1Language()
    If clsStation.Language = English Then
        Toolbar1.Buttons(1).Caption = "First"
        Toolbar1.Buttons(2).Caption = "Pre"
        Toolbar1.Buttons(3).Caption = "Next"
        Toolbar1.Buttons(4).Caption = "Last"
        Toolbar1.Buttons(5).Caption = ""
        Toolbar1.Buttons(6).Caption = "Add"
        Toolbar1.Buttons(7).Caption = "Edit"
        Toolbar1.Buttons(8).Caption = "Apply"
        Toolbar1.Buttons(9).Caption = "Cancel"
        Toolbar1.Buttons(10).Caption = "Delete"
        Toolbar1.Buttons(11).Caption = ""
        Toolbar1.Buttons(12).Caption = "Help"
        Toolbar1.Buttons(13).Caption = "Search"
        Toolbar1.Buttons(14).Caption = ""
        Toolbar1.Buttons(15).Caption = "Print"
        Toolbar1.Buttons(16).Caption = "Scan"
        Toolbar1.Buttons(17).Caption = ""
        Toolbar1.Buttons(18).Caption = "Refund"
        Toolbar1.Buttons(19).Caption = ""
        Toolbar1.Buttons(20).Caption = "Font"
        Toolbar1.Buttons(21).Caption = "Color"
        Toolbar1.Buttons(22).Caption = ""
        Toolbar1.Buttons(23).Caption = "Lang"
        Toolbar1.Buttons(24).Caption = "Phone"
        Toolbar1.Buttons(25).Caption = "KB"
        Toolbar1.Buttons(26).Caption = "Calc"
        Toolbar1.Buttons(27).Caption = ""
        Toolbar1.Buttons(28).Caption = "log out"
        Toolbar1.Buttons(29).Caption = "Exit"
    Else
        Toolbar1.Buttons(1).Caption = "«» œ«"
        Toolbar1.Buttons(2).Caption = "ﬁ»·Ì"
        Toolbar1.Buttons(3).Caption = "»⁄œÌ"
        Toolbar1.Buttons(4).Caption = "«‰ Â«"
        Toolbar1.Buttons(5).Caption = ""
        Toolbar1.Buttons(6).Caption = "«›“Êœ‰"
        Toolbar1.Buttons(7).Caption = "ÊÌ—«Ì‘"
        Toolbar1.Buttons(8).Caption = " À»  "
        Toolbar1.Buttons(9).Caption = "«‰’—«›"
        Toolbar1.Buttons(10).Caption = "Õ–›"
        Toolbar1.Buttons(11).Caption = ""
        Toolbar1.Buttons(12).Caption = "—«Â‰„«"
        Toolbar1.Buttons(13).Caption = "Ã” ÃÊ"
        Toolbar1.Buttons(14).Caption = ""
        Toolbar1.Buttons(15).Caption = "À»  Ê ç«Å"
        Toolbar1.Buttons(16).Caption = "«”ò‰"
        Toolbar1.Buttons(17).Caption = ""
        Toolbar1.Buttons(18).Caption = "„—ÃÊ⁄"
        Toolbar1.Buttons(19).Caption = ""
        Toolbar1.Buttons(20).Caption = "›Ê‰ "
        Toolbar1.Buttons(21).Caption = "—‰ê"
        Toolbar1.Buttons(22).Caption = ""
        Toolbar1.Buttons(23).Caption = "“»«‰"
        Toolbar1.Buttons(24).Caption = " ·›‰"
        Toolbar1.Buttons(25).Caption = "òÌ»Ê—œ"
        Toolbar1.Buttons(26).Caption = "„. Õ”«»"
        Toolbar1.Buttons(27).Caption = ""
        Toolbar1.Buttons(28).Caption = " —ò. ’"
        Toolbar1.Buttons(29).Caption = "Œ—ÊÃ"
    End If
'    If clsInvoiceValue.LanguageIcon = 0 Then Toolbar1.Buttons(23).Visible = False
'    If clsInvoiceValue.KeyboardIcon = 0 Then Toolbar1.Buttons(25).Visible = False
'    If clsInvoiceValue.ColorIcon = 0 Then Toolbar1.Buttons(21).Visible = False
'    If clsInvoiceValue.TelephoneIcon = 0 Then Toolbar1.Buttons(24).Visible = False

End Sub
Public Property Set FileCls(myVar As Object)
End Property


Private Sub tmrScreenSaver_Timer()
    timeInterval = timeInterval - 1
    If (timeInterval = 0) Then
        CallScreenSaver
    End If
End Sub

Private Sub tmrUdp_Timer()
    
'    If WinsockUdp.State = 1 Then
''        WinsockUdp.Close
''        modsock.ConnectToClient
'     Else
'        modsock.ConnectToClient
'     End If

End Sub
Public Sub CallScreenSaver()
    SendMessage Me.hWnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&
    AccessAfterClosingcash = True
    Unload frmAccess
    frmAccess.lblTitle.Caption = " —„“ ŒÊœ —« Ê«—œ ﬂ‰Ìœ"
    frmAccess.Show vbModal
    AccessAfterClosingcash = False
    timeInterval = clsInvoiceValue.ScreenSaverTime
End Sub
Private Sub mnuEscape_Click(index As Integer)
    CallScreenSaver
End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Call PresetScreenSaver

Dim varForm As Form
Dim frmact As Form
 
 On Error Resume Next
 
 Unload frmfactor
 Unload frmInput
 Unload frmMsg
 Unload frmDisMsg
For Each varForm In Forms
    If VarActForm = varForm.Name Or VarActForm = "" Then
        Set frmact = varForm
        Exit For
    End If
Next


Select Case Button.index
    Case 1:     'Home
            If Not IsNull(frmact) Then
                frmact.BeforeFirstKey
            End If
            If Not IsNull(frmact) Then
                frmact.FirstKey
            End If
    Case 2:     'PageDown
            If Not IsNull(frmact) Then
                frmact.BeforePreviousKey
            End If
            If Not IsNull(frmact) Then
                frmact.PreviousKey
            End If
    Case 3:     'PageUp Key
            If Not IsNull(frmact) Then
                frmact.BeforeNextKey
            End If
            If Not IsNull(frmact) Then
                frmact.NextKey
            End If
    Case 4:     'End Key
            If Not IsNull(frmact) Then
                frmact.BeforeLastKey
            End If
            If Not IsNull(frmact) Then
                frmact.LastKey
            End If
    Case 6:     'Add Key
            If Not IsNull(frmact) Then
                frmact.BeforeAdd
            End If

            If Not IsNull(frmact) Then
                frmact.MyFormAddEditMode = AddMode
                frmact.Add
            End If
    Case 7:     'Edit Key
            If Not IsNull(frmact) Then
                frmact.BeforeEdit
            End If
            
            If Not IsNull(frmact) Then
                frmact.Edit
            End If

    Case 8:     'Enter Key
            If Not IsNull(frmact) Then
                frmact.BeforeUpdate
            End If
            If Not IsNull(frmact) Then
                frmact.Update
            End If
                    
    Case 9:     'Esc Key (Cancel)
            If Not IsNull(frmact) Then
                frmact.BeforCancel
                frmact.Cancel
                frmact.AfterCancel
            End If
    
    Case 10:    'Delete Key
    
            If Not IsNull(frmact) Then
                frmact.Delete
            End If
                           
    Case 12:    'Help Key
'            Shell App.Path & "\Help\Help.htm"
            
            IsHelp = True
            Unload frmBrowser
            frmBrowser.Show
            IsHelp = False
            Call OnTopMe(frmBrowser, True)
''''            If Not IsNull(frmAct) Then
''''                frmAct.Help
''''            End If
    Case 13:    'Search (Find)
            If Not IsNull(frmact) Then
                frmact.Find
            End If
    Case 15:    'Printing
            If Not IsNull(frmact) Then
                frmact.Printing
            End If
    Case 16:    'Scan
            If Not IsNull(frmact) Then
                frmact.Scan
            End If
    Case 18    'Recursive
            If Not IsNull(frmact) Then
                frmact.UndoRedo
            End If
    'Case 19:
    '            If Not IsNull(frmAct) Then
    '                frmAct.Redo
    '            End If
    Case 20    'Font
            If Not IsNull(frmact) Then
                    frmFont.Show vbModal
            End If
    Case 21    'Color
            If Not IsNull(frmact) Then
                frmColor.Show vbModal
            End If
                
    Case 23
        If clsStation.Language = Farsi Then
            clsStation.Language = English
            SetKbLayout LANG_EN_US
            mdifrm.Toolbar1.Buttons(23).Value = tbrPressed
            ChangeToolBar1Language
            ChangeToolbar3Language
        Else
            clsStation.Language = Farsi
            SetKbLayout LANG_Pr_IR
            mdifrm.Toolbar1.Buttons(23).Value = tbrUnpressed
            ChangeToolBar1Language
            ChangeToolbar3Language
        End If
        SetStationSettingFile
        frmact.ChangeLanguage
    Case 24
        If Me.ActiveForm Is frmPhoneBook Then Exit Sub
''''        Set frmPhoneBook.PreviousForm = Me.ActiveForm
''''        Dim obj As Object
''''        For Each obj In Forms
''''            If TypeOf obj Is Form Then
''''                If LCase(obj.Name) <> "mdifrm" Then
''''                    Unload obj
''''                End If
''''            End If
''''
''''        Next obj
        If clsArya.PhoneBook = True Then
            frmPhoneBook.Show vbModal
            frmPhoneBook.SetFocus
        Else
            frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            frmact.Show
        End If
    Case 25
        Shell App.Path & "\Tools\osk.exe"
        
    Case 26
        Shell App.Path & "\Tools\calc.exe"
        
    Case 28
        If intVersion = Min Then
            ShowDisMessage "«„ﬂ«‰ «” ›«œÂ «“ «Ì‰ ﬁ«»·Ì  œ— ‰”ŒÂ Â«Ì »«·« — ÊÃÊœ œ«—œ", 1500
            Unload Me
            Exit Sub
        End If
        CallScreenSaver
        
    Case 29    'Exit Form
            If Not IsNull(frmact) And frmact.Name <> "mdifrm" Then
                frmact.ExitForm
            ElseIf frmact.Name = "mdifrm" Then
                ShowMessage "¬Ì« »—«Ì Œ—ÊÃ «“ »—‰«„Â «ÿ„Ì‰«‰ œ«—Ìœø", True, True, "»·Ì", "ŒÌ—"
                If mvarMsgIdx = vbYes Then
                    'the following lines is put in MDIFrm_Unload event
'                    ReDim Parameter(2) As Parameter
'                    Parameter(0) = GenerateInputParameter("@intUserNo", adInteger, 4, mvarCurUserNo)
'                    Parameter(1) = GenerateInputParameter("@intActionUserNo", adInteger, 4, 2)
'                    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
'                    RunParametricStoredProcedure "Insert_tblTotal_UserHistory", Parameter
                
                    Unload Me
                End If
            End If
End Select

End Sub


Private Sub mnuSelColor_Click()
''frmSelColor.Show
'''mdifrm.Arrange 0
End Sub

Private Sub mnuSelFont_Click()
''frmSelFont.Show
'''mdifrm.Arrange 0
End Sub

Private Sub Winsock_Farabin_Connect()
    bolSockIsConnected = True
    If mdifrm.Winsock_Farabin.State = sckConnected Then Winsock_Farabin.SendData strFarabin
End Sub

Private Sub Winsock_Farabin_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'    ShowDisMessage Description, 2000
    Winsock_Farabin.Close
End Sub

Private Sub Winsock_Farabin_SendComplete()
    Winsock_Farabin.Close
'    Winsock_Farabin.Connect
End Sub

Private Sub Winsock1_Connect()
    bolSockIsConnected = True
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   If Number = 10053 Or Number = 10061 Then
     frmMsg.fwlblMsg.Caption = "¬Ì«« ’«· œÊ»«—Â »Â ”—ÊÌ” »—ﬁ—«— ‘Êœø"
'     frmMsg.fwBtn(0).Visible = True
      frmMsg.fwBtn(0).ButtonType = flwButtonOk
      frmMsg.fwBtn(1).ButtonType = flwButtonCancel
      frmMsg.fwBtn(0).Caption = "»·Ì"
      frmMsg.fwBtn(1).Caption = "ŒÌ—"
      If LCase(VarActForm) = "frmlogin" Then
          FrmLogin.Command1(0).Enabled = False
      End If
            
      frmMsg.Show vbModal
      If mvarMsgIdx = vbYes Then
          modsock.ConnectSock
      Else
          Unload Me
      End If
   Else
      MsgBox Description
   End If
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Winsock1.GetData strSockRecive, vbString
    If strSockRecive = EOS Then
      Winsock1.Close
    End If
End Sub

Private Sub WinsockUdp_DataArrival(ByVal bytesTotal As Long)
    
    If LCase(VarActForm) = "frminvoice" Then
        On Error GoTo Err_Handler
        Dim index, LineNumber As Integer
        Dim StrCaller_Id As String
        WinsockUdp.GetData Msg, vbString
        If WinsockUdp.State = 1 Then WinsockUdp.Close
        If Len(LTrim(Msg)) > 4 Then
            LineNumber = Val(Left(LTrim(Msg), 1))
            index = Val(Mid(LTrim(Msg), 3, 1))
            StrCaller_Id = Mid(LTrim(Msg), 4)
            frmInvoice.GetCallerInfo -1, StrCaller_Id, LineNumber
        ElseIf Mid(LTrim(Msg), 2, 1) <> clsArya.StationNo And Val(Right(LTrim(Msg), 1)) <> clsArya.StationNo Then
    '        LineNumber = Val(Left(LTrim(MSG), 1))
    '        Index = Val(Mid(LTrim(MSG), 2, 1))
            index = Val(Left(LTrim(Msg), 1))
            frmInvoice.SetFWModemSetting (index)
            modsock.ConnectToClient
            If Len(LTrim(Msg)) <> 4 Then
                 WinsockUdp.SendData LTrim(Msg) & clsArya.StationNo
            End If
        End If
    End If
    If WinsockUdp.State = 0 Then modsock.ConnectToClient
Exit Sub
Err_Handler:
    LogSaveNew "mdifrm => ", err.Description, err.Number, err.Source, "WinsockUdp_DataArrival"
    'ShowDisMessageNoModal err.Description, 1000
    err.Clear
    If WinsockUdp.State = 0 Then modsock.ConnectToClient

End Sub

Public Sub AutoBackup()
    On Error GoTo ErrHandler
    Dim cnn As New ADODB.Connection
    cnn.ConnectionString = strConnectionString
    cnn.Open
    If Not filetemp.FolderExists(App.Path & "\BackUp") Then
        filetemp.CreateFolder (App.Path & "\BackUp")
    End If
    With cmd
        .ActiveConnection = cnn
        .CommandType = adCmdText
        
'                .CommandText = "USE master" & _
'                               " EXEC sp_addumpdevice 'disk', 'tmpTotal', '" & Dir1.Path & "\" & Mid(txtFile.Text, 1, 6) & "_" & CStr(cnt) & ".Bak" & " ' " & _
'                               " BACKUP DATABASE " & clsArya.DbName & " To tmpTotal " & _
'                               " exec sp_dropdevice 'tmpTotal' "
        .CommandText = " BACKUP DATABASE [" & clsArya.DbName & "] To Disk = N'" & App.Path & "\BackUp\" & CurrentDateNumber & ".Bak' WITH  INIT ,  NOUNLOAD ,   NAME = N'" & clsArya.DbName & " backup',  NOSKIP ,  STATS = 10,  NOFORMAT"
    
    End With
    cmd.Execute
    cmd.Cancel
    ShowDisMessageNoModal "«ÌÃ«œ ê—œÌœ" & App.Path & "\BackUp\" & "‰”ŒÂ Å‘ Ì»«‰ »Â  «—ÌŒ —Ê“ Ã«—Ì œ— ", 3000
    
    frmDisMsg.Timer1.Interval = 1000
    frmDisMsg.lblMessage = "”Ì” „ œ— Õ«· Å«ﬂ”«“Ì ‰”ŒÂ Â«Ì Å‘ Ì»«‰ ﬁœÌ„Ì „Ì »«‘œ "
    frmDisMsg.lblMessage = frmDisMsg.lblMessage & vbLf & "·ÿ›« „‰ Ÿ— »„«‰Ìœ "
    frmDisMsg.Show
    DoEvents
    On Error Resume Next
    DeleteOldBackup
    On Error GoTo 0
    
Exit Sub
ErrHandler:
    MsgBox "œ— ê—› ‰ ‰”ŒÂ Å‘ Ì»«‰ „‘ò· ÊÃÊœ œ«—œ "
End Sub

Private Sub DeleteOldBackup()

'Delete all SQL Server backup files more than 5 days old

Dim oFS, oSQLBackupFol, oFol, oFil
Dim sPattern
sPattern = "*.bak"
Set oFS = CreateObject("Scripting.FileSystemObject")
'Set oSQLBackupFol = oFS.GetFolder("drive:\path")   'Change this as appropriate
Set oSQLBackupFol = oFS.GetFolder(App.Path & "\BackUp")     'Change this as appropriate

'For Each oFol In oSQLBackupFol.SubFolders    'get subfolders of the above path
    For Each oFil In oSQLBackupFol.Files   'get each file in subfolder
        If oFil.DateCreated < Now - 14 Then 'Change this as appropriate
            If UCase(Right(oFil.Name, 4)) = ".BAK" Then
                oFil.Delete
            End If
        End If
    Next
'Next
ShowDisMessageNoModal "‰”ŒÂ Â«Ì Å‘ Ì»«‰ ﬁœÌ„Ì  — «“ 2 Â› Â ﬁ»· Å«ﬂ”«“Ì ê—œÌœ  ", 2000

Set oFS = Nothing
Set oSQLBackupFol = Nothing
    
Exit Sub
ErrHandler:
    MsgBox "œ— Å«ﬂ”«“Ì ‰”ŒÂ Â«Ì Å‘ Ì»«‰ ﬁœÌ„Ì „‘ò· ÊÃÊœ œ«—œ "
End Sub


