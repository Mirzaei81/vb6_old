VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmReports 
   BackColor       =   &H00FF8080&
   ClientHeight    =   8655
   ClientLeft      =   105
   ClientTop       =   495
   ClientWidth     =   13800
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmReports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   13800
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   10
      Tab             =   9
      TabsPerRow      =   5
      TabHeight       =   882
      WordWrap        =   0   'False
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ê“«—‘ ò«·« Ê €—›Â"
      TabPicture(0)   =   "frmReports.frx":A4C2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fwlblRep(42)"
      Tab(0).Control(1)=   "fwlblRep(41)"
      Tab(0).Control(2)=   "fwlblRep(40)"
      Tab(0).Control(3)=   "fwlblRep(0)"
      Tab(0).Control(4)=   "fwlblRep(2)"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "ê“«—‘ ›—Ê‘"
      TabPicture(1)   =   "frmReports.frx":A4DE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fwlblRep(26)"
      Tab(1).Control(1)=   "fwlblRep(32)"
      Tab(1).Control(2)=   "fwlblRep(19)"
      Tab(1).Control(3)=   "fwlblRep(18)"
      Tab(1).Control(4)=   "fwlblRep(11)"
      Tab(1).Control(5)=   "fwlblRep(6)"
      Tab(1).Control(6)=   "fwlblRep(5)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "«‰»«— Ê ò«·«"
      TabPicture(2)   =   "frmReports.frx":A4FA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fwlblRep(21)"
      Tab(2).Control(1)=   "fwlblRep(28)"
      Tab(2).Control(2)=   "fwlblRep(45)"
      Tab(2).Control(3)=   "fwlblRep(44)"
      Tab(2).Control(4)=   "fwlblRep(43)"
      Tab(2).Control(5)=   "fwlblRep(27)"
      Tab(2).Control(6)=   "fwlblRep(25)"
      Tab(2).Control(7)=   "fwlblRep(24)"
      Tab(2).Control(8)=   "fwlblRep(23)"
      Tab(2).Control(9)=   "fwlblRep(22)"
      Tab(2).Control(10)=   "fwlblRep(16)"
      Tab(2).Control(11)=   "fwlblRep(15)"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "ê“«—‘ «”‰«œ"
      TabPicture(3)   =   "frmReports.frx":A516
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fwlblRep(38)"
      Tab(3).Control(1)=   "fwlblRep(37)"
      Tab(3).Control(2)=   "fwlblRep(36)"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "„—«”„"
      TabPicture(4)   =   "frmReports.frx":A532
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fwlblRep(33)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "ÅÌò  Ê ê«—”Ê‰"
      TabPicture(5)   =   "frmReports.frx":A54E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fwlblRep(17)"
      Tab(5).Control(1)=   "fwlblRep(3)"
      Tab(5).Control(2)=   "fwlblRep(4)"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "ê“«—‘ Œ—Ìœ"
      TabPicture(6)   =   "frmReports.frx":A56A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fwlblRep(29)"
      Tab(6).Control(1)=   "fwlblRep(13)"
      Tab(6).Control(2)=   "fwlblRep(12)"
      Tab(6).Control(3)=   "fwlblRep(14)"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "„‘ —òÌ‰ Ê  «„Ì‰ ò‰‰œê«‰"
      TabPicture(7)   =   "frmReports.frx":A586
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fwlblRep(8)"
      Tab(7).Control(1)=   "fwlblRep(7)"
      Tab(7).Control(2)=   "fwlblRep(39)"
      Tab(7).Control(3)=   "fwlblRep(31)"
      Tab(7).Control(4)=   "fwlblRep(30)"
      Tab(7).Control(5)=   "fwlblRep(20)"
      Tab(7).Control(6)=   "fwlblRep(47)"
      Tab(7).ControlCount=   7
      TabCaption(8)   =   "ê“«—‘ Ã«Ì“Â"
      TabPicture(8)   =   "frmReports.frx":A5A2
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "fwlblRep(35)"
      Tab(8).Control(1)=   "fwlblRep(34)"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "êÊ‰«êÊ‰"
      TabPicture(9)   =   "frmReports.frx":A5BE
      Tab(9).ControlEnabled=   -1  'True
      Tab(9).Control(0)=   "fwlblRep(46)"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "fwlblRep(10)"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).Control(2)=   "fwlblRep(9)"
      Tab(9).Control(2).Enabled=   0   'False
      Tab(9).Control(3)=   "fwlblRep(1)"
      Tab(9).Control(3).Enabled=   0   'False
      Tab(9).ControlCount=   4
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   2
         Left            =   -67080
         Tag             =   "RepDetailGoodsSale"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ›—Ê‘ —Ì“ ò«·« Â« Ê  Œ›Ì›«  ﬂ«·«Â«"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A5DA
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   0
         Left            =   -64560
         Tag             =   "RepGroupGoodsSale"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ›—Ê‘ ê—ÊÂÌ ò«·« Â«"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A5F6
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   40
         Left            =   -69480
         Tag             =   "RepCustPricDiscount"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ›—Ê‘       €—›Â Â« "
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9
         Alignment       =   2
         Picture         =   "frmReports.frx":A612
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   41
         Left            =   -71880
         Tag             =   "RepCustPricDiscount"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ›—Ê‘         «Ì” ê«Â« »—«”«” €—›Â Â«"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9
         Alignment       =   2
         Picture         =   "frmReports.frx":A62E
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   42
         Left            =   -74280
         Tag             =   "RepBranchSale_CrossTab"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ›—Ê‘         ‘⁄»Â Â« »—«”«” €—›Â Â«"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9
         Alignment       =   2
         Picture         =   "frmReports.frx":A64A
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   5
         Left            =   -64560
         Tag             =   "RepFichSale"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ›—Ê‘ »— «”«” ›Ì‘"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A666
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   6
         Left            =   -71880
         Tag             =   "RepServeKindSell"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ «‰Ê«⁄ ›—Ê‘"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A682
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   11
         Left            =   -66000
         Tag             =   $"frmReports.frx":A69E
         Top             =   4320
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ »— «”«” ‘Ì› "
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A6AC
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   18
         Left            =   -67560
         Tag             =   "RepCashInvoice"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ›—Ê‘ ’‰œÊﬁ"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A6C8
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   19
         Left            =   -70560
         Tag             =   "RepStationSaleSummery"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ Œ·«’Â ›—Ê‘ ’‰œÊﬁ"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A6E4
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   15
         Left            =   -65040
         Tag             =   "RepMojodi"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ „ÊÃÊœÌ  ⁄œ«œÌ «‰»«— - œ—  «—ÌŒ Œ«’"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A700
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   16
         Left            =   -67680
         Tag             =   "RepMojodi"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘  „ÊÃÊœÌ     —Ì«·Ì «‰»«— - œ—  «—ÌŒ Œ«’"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A71C
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   22
         Left            =   -65040
         Tag             =   "RepGoodList"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "·Ì”  ò«·« Â«Ì  ⁄—Ì› ‘œÂ"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A738
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   23
         Left            =   -67560
         Tag             =   $"frmReports.frx":A754
         Top             =   3960
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ·Ì”   €ÌÌ—« "
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A76C
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   24
         Left            =   -70320
         Tag             =   "RepUsedGoodAmount"
         Top             =   5880
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ „’—›    ﬂ«·« Ê „Ê«œ «Ê·ÌÂ"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A788
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   25
         Left            =   -70320
         Tag             =   "RepLossGoodAmount"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ÷«Ì⁄« "
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A7A4
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   32
         Left            =   -68880
         Top             =   4320
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ›—Ê‘     „Ì“Â«"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A7C0
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   26
         Left            =   -73440
         Tag             =   "RepStationSaleSummaryByUser"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ Œ·«’Â ›—Ê‘ ’‰œﬁœ«—«‰"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A7DC
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   4
         Left            =   -69240
         Tag             =   "RepBedehDetail"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ »œÂﬂ«—Ì ÅÌò Ã“∆Ì« "
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A7F8
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   3
         Left            =   -65400
         Tag             =   "RepBedeh"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ »œÂﬂ«—Ì ÅÌﬂ"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A814
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   17
         Left            =   -72840
         Tag             =   "RepGarson"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ »œÂò«—Ì ê«—”Ê‰ "
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A830
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   27
         Left            =   -73080
         Tag             =   "RepGetOrderGoodAmount"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ «“ ﬂ«·«Â«Ì »Â ‰ﬁÿÂ ”›«—‘ —”ÌœÂ"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A84C
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   43
         Left            =   -67440
         Tag             =   "RepTurnRecipt"
         Top             =   5880
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ê—œ‘       —”Ìœ Â«"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A868
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   44
         Left            =   -70320
         Tag             =   "RepMojodi"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ⁄„·ò—œ  ⁄œ«œÌ «‰»«—  œ—”«· „«·Ì"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A884
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   45
         Left            =   -73080
         Tag             =   "RepMojodi"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘  ⁄„·ò—œ —Ì«·Ì «‰»«— - œ— ”«· „«·Ì"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A8A0
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   8
         Left            =   -68640
         Tag             =   "RepCustPrice"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "·Ì”  Œ—Ìœ      „‘ —ﬂÌ‰ Ê  «„Ì‰ ﬂ‰‰œê«‰ -Ã“∆Ì« "
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A8BC
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   7
         Left            =   -65160
         Tag             =   "RepCustPricDiscount"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "·Ì”  Œ—Ìœ „‘ —Ì«‰ Ê  «„Ì‰ ﬂ‰‰œê«‰"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A8D8
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   39
         Left            =   -71760
         Tag             =   "RepSaleShiftDailyPrize"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ’Ê— Õ”«» „‘ —Ì«‰"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A8F4
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   31
         Left            =   -69840
         Tag             =   "RepCustPriceReturn"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "»—ê‘  «“ Œ—Ìœ      „‘ —ﬂÌ‰ Ê  «„Ì‰ ﬂ‰‰œê«‰ -Ã“∆Ì« "
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A910
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   30
         Left            =   -66720
         Tag             =   "RepCustPricDiscountReturn"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "»—ê‘  «“Œ—Ìœ „‘ —Ì«‰ Ê  «„Ì‰ ﬂ‰‰œê«‰"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A92C
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   28
         Left            =   -65040
         Tag             =   "RepDetailGoodsSaleReturn"
         Top             =   5760
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ »—ê‘  «“ ›—Ê‘ —Ì“ ﬂ«·«Â«"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A948
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   36
         Left            =   -65880
         Top             =   2640
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ›—Ê‘ «ﬁ”«ÿÌ "
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A964
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   37
         Left            =   -69240
         Tag             =   "RepSaleShiftDailyPrize"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ »‰ Â«Ì œ—Ì«› Ì"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A980
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   38
         Left            =   -72360
         Tag             =   "RepSaleShiftDailyPrize"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ çﬂ Â«Ì œ—Ì«› Ì"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A99C
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   20
         Left            =   -73320
         Tag             =   "RepCustomerList"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "·Ì”  „‘ —òÌ‰"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A9B8
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   14
         Left            =   -65040
         Tag             =   "RepGroupGoodsBuy"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ Œ—Ìœ       ê—ÊÂÌ ò«·« Â«"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":A9D4
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   12
         Left            =   -70680
         Tag             =   $"frmReports.frx":A9F0
         Top             =   2760
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ Œ—Ìœ »— «”«” ›Ì‘"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":AA00
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   13
         Left            =   -67920
         Tag             =   "RepDetailGoodsBuy"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ Œ—Ìœ —Ì“ ﬂ«·«Â«"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":AA1C
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   29
         Left            =   -73320
         Tag             =   "RepDetailGoodsBuyReturn"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ »—ê‘  «“ Œ—Ìœ —Ì“ ﬂ«·«Â«"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":AA38
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   34
         Left            =   -67440
         Tag             =   "RepDailyPrize"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ Ã«Ì“Â      »—«”«” —Ê“"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":AA54
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   35
         Left            =   -70560
         Tag             =   "RepSaleShiftDailyPrize"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ›—Ê‘ ‘Ì›  »« Ã«Ì“Â  »—«”«” —Ê“"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":AA70
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   1
         Left            =   10080
         Tag             =   $"frmReports.frx":AA8C
         Top             =   2880
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ›—Ê‘ ”«⁄ Ì/œ—’œÌ"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":AAA0
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   9
         Left            =   4320
         Tag             =   "RepSystemGroup"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ê—ÊÂÂ«Ì ”Ì” „Ì"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":AABC
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   10
         Left            =   7200
         Tag             =   "RepSerialFich"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ »— «”«” ‘„«—Â ”—Ì«·"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":AAD8
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   46
         Left            =   1440
         Tag             =   "RepAdditionalService"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ”—ÊÌ” Â«Ì œ—Ì«› Ì «“ „‘ —Ì«‰"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":AAF4
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   33
         Left            =   -69600
         Tag             =   "RepDailyWeeding"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   " ê“«—‘ —“—Ê „—«”„  »—«”«” —Ê“      "
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":AB10
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   21
         Left            =   -73080
         Tag             =   "RepInventoryRecipt"
         Top             =   5880
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   " »— «”«” ‰Ê⁄ ›«ò Ê—"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":AB2C
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwlblRep 
         Height          =   975
         Index           =   47
         Left            =   -74400
         Tag             =   "RepSaleShiftDailyPrize"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1720
         Enabled         =   -1  'True
         Caption         =   "ê“«—‘ ’Ê— Õ”«» „‘ —Ì«‰ »« „«‰œÂ «“ ﬁ»·"
         FirstColor      =   16761024
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmReports.frx":AB48
         BorderStyle     =   1
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2160
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   480
      Top             =   120
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmReports.frx":AB64
      TabIndex        =   1
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "·Ì”  ê“«—‘« "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CounterRep As Integer
Public IndexRep As Integer
Dim Parameter() As Parameter
Dim ArrayUbound As Integer
Private clsDate As New clsDate
Dim i As Integer
Dim accessform As Boolean

Private Sub Form_Activate()
    mdifrm.Toolbar3.Visible = False
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

    If mdifrm.clsFormAccess.frmReports = False Then
        Unload Me
        Exit Sub
    End If

    CenterTop Me
    CounterRep = 0
    
    Dim obj As Object
    For Each obj In Forms
        If TypeOf obj Is Form Then
            If obj.Name <> "mdifrm" And obj.Name <> Me.Name Then
                obj.Hide
            End If
        End If

    Next obj

    Dim Rst As New ADODB.Recordset
    
    If Rst.State <> 0 Then Rst.Close
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@UserId", adInteger, 4, mvarCurUserNo)
    Parameter(1) = GenerateInputParameter("@intObjectType", adInteger, 4, 2)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set Rst = RunParametricStoredProcedure2Rec("GetUserAccess", Parameter)
        
    For i = 0 To Me.fwlblRep.Count - 1
        If i = 15 Or i = 16 Then
            Me.fwlblRep(i).Visible = False
        Else
            Me.fwlblRep(i).Visible = True
            Me.fwlblRep(i).Enabled = False
        End If
    Next i
Dim Tempa As String
Dim Tempb As String
    If Not (Rst.EOF = True And Rst.BOF = True) Then
            While Rst.EOF <> True
                For i = 0 To Me.fwlblRep.Count - 1
                    Tempa = LCase(LTrim(RTrim((Me.fwlblRep(i).Tag))))
                    Tempb = LCase(LTrim(RTrim((Rst.Fields("ObjectId").Value))))
                    
                    If InStr(1, Tempa, Tempb, vbTextCompare) Then
                        Me.fwlblRep(i).Enabled = True
                        Exit For
                    End If
                Next i
                Rst.MoveNext
                
            Wend
    End If

Set Rst = Nothing
If clsArya.Delivery = False Then
    Me.fwlblRep(3).Enabled = False
  '  Me.fwlblRep(3).Caption = ""
    Me.fwlblRep(4).Enabled = False
  '  Me.fwlblRep(4).Caption = ""
End If
If (Val(strCategory) >= 0 And Val(strCategory) <= 7) Then
    Me.fwlblRep(32).Enabled = True
Else
    Me.fwlblRep(6).Enabled = False
  '  Me.fwlblRep(6).Caption = ""
    Me.fwlblRep(17).Enabled = False
  '  Me.fwlblRep(17).Caption = ""
    Me.fwlblRep(23).Enabled = False
  '  Me.fwlblRep(23).Caption = ""
    Me.fwlblRep(32).Enabled = False
  '  Me.fwlblRep(32).Caption = ""
End If
If clsArya.ProductSystem = False Then
    Me.fwlblRep(24).Enabled = False
  '  Me.fwlblRep(24).Caption = ""
    Me.fwlblRep(25).Enabled = False
  '  Me.fwlblRep(25).Caption = ""
End If
If clsArya.TableGarson = False Then
    Me.fwlblRep(17).Enabled = False
  '  Me.fwlblRep(17).Caption = ""
    Me.fwlblRep(32).Enabled = False
  '  Me.fwlblRep(32).Caption = ""
End If
If clsArya.Customers = False Then
    Me.fwlblRep(20).Enabled = False
  '  Me.fwlblRep(20).Caption = ""
    Me.fwlblRep(39).Enabled = False
  '  Me.fwlblRep(39).Caption = ""
    Me.fwlblRep(47).Enabled = False
End If
If clsArya.Customers = False And clsArya.StoreGroup = False Then
    Me.fwlblRep(7).Enabled = False
  '  Me.fwlblRep(7).Caption = ""
    Me.fwlblRep(8).Enabled = False
  '  Me.fwlblRep(8).Caption = ""
    Me.fwlblRep(30).Enabled = False
  '  Me.fwlblRep(30).Caption = ""
    Me.fwlblRep(31).Enabled = False
  '  Me.fwlblRep(31).Caption = ""
End If
If clsArya.StoreGroup = False Then
    Me.fwlblRep(12).Enabled = False
  '  Me.fwlblRep(12).Caption = ""
    Me.fwlblRep(13).Enabled = False
  '  Me.fwlblRep(13).Caption = ""
    Me.fwlblRep(14).Enabled = False
  '  Me.fwlblRep(14).Caption = ""
    Me.fwlblRep(15).Enabled = False
  '  Me.fwlblRep(15).Caption = ""
    Me.fwlblRep(16).Enabled = False
  '  Me.fwlblRep(16).Caption = ""
    Me.fwlblRep(27).Enabled = False
  '  Me.fwlblRep(27).Caption = ""
    Me.fwlblRep(29).Enabled = False
  '  Me.fwlblRep(29).Caption = ""
   Me.fwlblRep(44).Enabled = False
  '  Me.fwlblRep(44).Caption = ""
    Me.fwlblRep(45).Enabled = False
  '  Me.fwlblRep(45).Caption = ""
End If
Me.fwlblRep(36).Enabled = True
Me.fwlblRep(37).Enabled = True
Me.fwlblRep(38).Enabled = True
Me.fwlblRep(39).Enabled = True
Me.fwlblRep(46).Enabled = True
Me.fwlblRep(47).Enabled = True
If Me.fwlblRep(15).Enabled = True Then
    Me.fwlblRep(16).Enabled = True
    Me.fwlblRep(44).Enabled = True
    Me.fwlblRep(45).Enabled = True
End If
If Val(strCategory) = 24 Then
    Me.fwlblRep(33).Enabled = False
    Me.fwlblRep(34).Enabled = False
    Me.fwlblRep(40).Enabled = True
    Me.fwlblRep(41).Enabled = True
    Me.fwlblRep(42).Enabled = True
    Me.fwlblRep(43).Enabled = True
End If

    formloadFlag = False
    Me.Left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 0 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 0 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.Top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    formloadFlag = True

    SSTab1.Tab = SstabIndex

End Sub


Public Sub ExitForm()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdifrm.Toolbar3.Visible = True
    VarActForm = ""
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    Dim obj As Object
    For Each obj In Forms
        If TypeOf obj Is Form Then
            If obj.Name <> "mdifrm" And obj.Name <> Me.Name Then
                obj.Show
            End If
        End If

    Next obj


    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

    SstabIndex = SSTab1.Tab

End Sub


Private Sub fwlblRep_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    IndexRep = Index
    Unload Me
    
    accessform = False
    
    Dim obj As Object
    For Each obj In Forms
        If TypeOf obj Is Form Then
            If obj.Name <> "mdifrm" Then
                obj.Hide
            End If
        End If

    Next obj
    
    Select Case Index
        Case 0
            CRepFlag = "RepGroupGoodsSale"
            If mdifrm.clsFormAccess.RepGroupGoodsSale Then
               accessform = True
            End If
        Case 1
            CRepFlag = "RepTimePerSell"
            If mdifrm.clsFormAccess.RepTimePerSell Then
               accessform = True
            End If
        Case 2
            CRepFlag = "RepDetailGoodsSale"
            If mdifrm.clsFormAccess.RepDetailGoodsSale Then
               accessform = True
            End If
        Case 3
            CRepFlag = "RepBedeh"
            If mdifrm.clsFormAccess.RepBedeh Then
               accessform = True
            End If
        Case 4
            CRepFlag = "RepBedehDetail"
            If mdifrm.clsFormAccess.RepBedehDetail Then
               accessform = True
            End If
        Case 5
            CRepFlag = "RepFichSale"
            If mdifrm.clsFormAccess.RepFichSale Then
               accessform = True
            End If
        Case 6
            CRepFlag = "RepServeKindSell"
            If mdifrm.clsFormAccess.RepServeKindSell Then
               accessform = True
            End If
        Case 7
            CRepFlag = "RepCustPricDiscount"
            If mdifrm.clsFormAccess.RepCustPricDiscount Then
               accessform = True
            End If
        Case 8
            CRepFlag = "RepCustPrice"
            If mdifrm.clsFormAccess.RepCustPrice Then
               accessform = True
            End If
        Case 9
            CRepFlag = "RepSystemGroup"
            If mdifrm.clsFormAccess.RepSystemGroup Then
               accessform = True
            End If
        Case 10
            CRepFlag = "RepSerialFich"
            If mdifrm.clsFormAccess.RepSerialFich Then
               accessform = True
            End If
        Case 11
            CRepFlag = "RepShift"
            If mdifrm.clsFormAccess.RepShift Then
               accessform = True
            End If
        Case 12
            CRepFlag = "RepFichBuy"
            If mdifrm.clsFormAccess.RepFichBuy Then
               accessform = True
            End If
        Case 13
            CRepFlag = "RepDetailGoodsBuy"
            If mdifrm.clsFormAccess.RepDetailGoodsBuy Then
               accessform = True
            End If
        Case 14
            CRepFlag = "RepGroupGoodsBuy"
            If mdifrm.clsFormAccess.RepGroupGoodsBuy Then
               accessform = True
            End If
        Case 15
            CRepFlag = "RepMojodi"
            If mdifrm.clsFormAccess.RepMojodi Then
               accessform = True
            End If
        Case 16
            CRepFlag = "RepMojodiRial"
            If mdifrm.clsFormAccess.RepMojodi Then
               accessform = True
            End If
        Case 17
            CRepFlag = "RepGarson"
            If mdifrm.clsFormAccess.RepGarson Then
               accessform = True
            End If
        Case 18
            CRepFlag = "RepCashInvoice"
            If mdifrm.clsFormAccess.RepCashInvoice Then
               accessform = True
            End If
        Case 19
            CRepFlag = "RepStationSaleSummery"
            If mdifrm.clsFormAccess.RepStationSaleSummery Then
               accessform = True
            End If
        Case 20
            CRepFlag = "RepCustomerList"
            If mdifrm.clsFormAccess.RepCustomerList Then
               accessform = True
            End If
        Case 21
            CRepFlag = "RepInventoryRecipt"
            If mdifrm.clsFormAccess.RepInventoryRecipt Then
               accessform = True
            End If
        Case 22
            CRepFlag = "RepGoodList"
            If mdifrm.clsFormAccess.RepGoodList Then
               accessform = True
            End If
        Case 23
            CRepFlag = "RepGoodDifferences"
            If mdifrm.clsFormAccess.RepGoodDifferences Then
               accessform = True
            End If
        Case 24
            CRepFlag = "RepUsedGoodAmount"
            If mdifrm.clsFormAccess.RepUsedGoodAmount Then
               accessform = True
            End If
        Case 25
            CRepFlag = "RepLossGoodAmount"
            If mdifrm.clsFormAccess.RepLossGoodAmount Then
               accessform = True
            End If
        Case 26
            CRepFlag = "RepStationSaleSummaryByUser"
            If mdifrm.clsFormAccess.RepStationSaleSummaryByUser Then
               accessform = True
            End If
        Case 27
            CRepFlag = "RepGetOrderGoodAmount"
            If mdifrm.clsFormAccess.RepGetOrderGoodAmount Then
               accessform = True
            End If
        Case 28
            CRepFlag = "RepDetailGoodsSaleReturn"
            'If mdifrm.clsFormAccess.RepDetailGoodsSaleReturn Then
               accessform = True
            'End If
        Case 29
            CRepFlag = "RepDetailGoodsBuyReturn"
            'If mdifrm.clsFormAccess.RepDetailGoodsBuyReturn Then
               accessform = True
            'End If
        Case 30
            CRepFlag = "RepCustPricDiscountReturn"
            'If mdifrm.clsFormAccess.RepCustPricDiscountReturn Then
               accessform = True
            'End If
        Case 31
            CRepFlag = "RepCustPriceReturn"
            'If mdifrm.clsFormAccess.RepCustPriceReturn Then
               accessform = True
            'End If
         Case 32
            CRepFlag = "RepTableSellDetail"
            If mdifrm.clsFormAccess.RepSaleShiftDailyPrize Then
               accessform = True
            End If
        Case 33
            CRepFlag = "RepDailyWeeding"
            If mdifrm.clsFormAccess.RepDailyWeeding Then
               accessform = True
            End If
        Case 34
            CRepFlag = "RepDailyPrize"
            If mdifrm.clsFormAccess.RepDailyPrize Then
               accessform = True
            End If
         Case 35
            CRepFlag = "RepSaleShiftDailyPrize"
            If mdifrm.clsFormAccess.RepSaleShiftDailyPrize Then
               accessform = True
            End If
        Case 36
            CRepFlag = "RepCustomerLoan"
            'If mdifrm.clsFormAccess.RepCustPriceReturn Then
               accessform = True
            'End If
        Case 37
            CRepFlag = "RepCredit"
            'If mdifrm.clsFormAccess.RepStationSale_CrossTab Then
               accessform = True
            'End If
        Case 38
            CRepFlag = "RepCheque"
            'If mdifrm.clsFormAccess.RepStationSale_CrossTab Then
               accessform = True
            'End If
        Case 39
            CRepFlag = "RepCustomerBillPayment"
            'If mdifrm.clsFormAccess.RepStationSale_CrossTab Then
               accessform = True
            'End If
        Case 40
            CRepFlag = "RepSubInventory"
            'If mdifrm.clsFormAccess.RepCustPriceReturn Then
               accessform = True
            'End If
        Case 41
            CRepFlag = "RepStationSale_CrossTab"
            'If mdifrm.clsFormAccess.RepStationSale_CrossTab Then
               accessform = True
            'End If
        Case 42
            CRepFlag = "RepBranchSale_CrossTab"
            'If mdifrm.clsFormAccess.RepBranchSale_CrossTab Then
               accessform = True
            'End If
        Case 43
            CRepFlag = "RepTurnRecipt"
            'If mdifrm.clsFormAccess.RepTurnRecipt Then
               accessform = True
           ' End If
        Case 44
            CRepFlag = "RepMojodiYear"
            If mdifrm.clsFormAccess.RepMojodi Then
               accessform = True
            End If
         Case 45
            CRepFlag = "RepMojodiRialYear"
            If mdifrm.clsFormAccess.RepMojodi Then
               accessform = True
            End If
        Case 46
            CRepFlag = "RepAdditionalServices"
'''            If mdifrm.clsFormAccess.RepMojodi Then
               accessform = True
''        End If
        Case 47
            CRepFlag = "RepCustomerBillPayment_Remain"
               'If mdifrm.clsFormAccess.RepStationSale_CrossTab Then
                accessform = True
         'End If
    End Select
    
    If accessform = True Then
        frmRep.Show
    Else
        frmDisMsg.lblMessage = " ‘„« »Â «Ì‰ ê“«—‘ œ” —”Ì ‰œ«—Ìœ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
       
    End If
    
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub


Private Sub SetFirstToolBar()

    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True

End Sub





