VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmbarcode 
   Caption         =   "                                                                               ç«Å »«—òœ     "
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmbarcode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11400
   Begin VB.CommandButton Cmdprint 
      BackColor       =   &H00008000&
      Caption         =   "Å—Ì‰ "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      TabIndex        =   39
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H000080FF&
      Cancel          =   -1  'True
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MaskColor       =   &H000000FF&
      TabIndex        =   38
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Frame Frame8 
      Height          =   6255
      Left            =   0
      TabIndex        =   20
      Top             =   240
      Width           =   5055
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0FF&
         Caption         =   " ⁄—Ì› ”—⁄  Ê òÌ›Ì  ç«Å"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   4920
         Width           =   4455
         Begin FLWCtrls.FWNumericTextBox FWNumericSpeed 
            Height          =   495
            Left            =   2400
            TabIndex        =   34
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            Max             =   15
            Min             =   1
            Value           =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FLWCtrls.FWNumericTextBox FWNumericDensity 
            Height          =   495
            Left            =   240
            TabIndex        =   35
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            Max             =   15
            Min             =   1
            Value           =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "òÌ›Ì "
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1080
            TabIndex        =   37
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "”—⁄ "
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3240
            TabIndex        =   36
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         Caption         =   " ⁄—Ì› «— ›«⁄ »«—ﬂœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   3600
         Width           =   4455
         Begin VB.TextBox TxtBarcodeHeight 
            Height          =   495
            Left            =   1200
            TabIndex        =   31
            Text            =   "6"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "mm-«— ›«⁄"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1920
            TabIndex        =   32
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         Caption         =   " ⁄—Ì› «»⁄«œ ·Ì»· "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1680
         Width           =   4455
         Begin VB.TextBox TxtHGap 
            Height          =   495
            Left            =   1320
            TabIndex        =   26
            Text            =   "2"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Txtheight 
            Height          =   495
            Left            =   2520
            TabIndex        =   25
            Text            =   "18"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox TxtWidth 
            Height          =   495
            Left            =   240
            TabIndex        =   24
            Text            =   "30"
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "mm-H Gap"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2040
            TabIndex        =   29
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "mm-⁄—÷"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   960
            TabIndex        =   28
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "mm-«— ›«⁄"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3240
            TabIndex        =   27
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   " ⁄—Ì› ·Ì»· Å—Ì‰ —"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   360
         Width           =   4455
         Begin VB.ComboBox CmbPrinter 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   3735
         End
      End
   End
   Begin VB.Frame Frame9 
      Height          =   6255
      Left            =   5160
      TabIndex        =   2
      Top             =   240
      Width           =   6015
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         Caption         =   " ⁄—Ì› ›«’·Â ç«Å «“ ﬂ‰«— ·Ì»·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   4560
         Width           =   5535
         Begin VB.TextBox TxtDistance 
            Height          =   495
            Left            =   1200
            TabIndex        =   18
            Text            =   "6"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "mm-⁄—÷"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1920
            TabIndex        =   19
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   " ⁄—Ì› «‰œ«“Â ›Ê‰  "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2760
         Width           =   5535
         Begin FLWCtrls.FWNumericTextBox FWNumericprice 
            Height          =   495
            Left            =   3000
            TabIndex        =   9
            Top             =   1080
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            Value           =   30
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FLWCtrls.FWNumericTextBox FWNumericbarcode 
            Height          =   495
            Left            =   240
            TabIndex        =   10
            Top             =   1080
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            Value           =   30
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FLWCtrls.FWNumericTextBox FWNumericgood 
            Height          =   495
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            Value           =   25
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FLWCtrls.FWNumericTextBox FWNumericstore 
            Height          =   495
            Left            =   3000
            TabIndex        =   12
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            Value           =   30
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„ "
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
            Left            =   3960
            TabIndex        =   16
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
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
            Height          =   495
            Left            =   1320
            TabIndex        =   15
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "‰«„ ò«·«"
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
            Left            =   1320
            TabIndex        =   14
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "›—Ê‘ê«Â"
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
            Left            =   3960
            TabIndex        =   13
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   2055
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   5535
         Begin VB.ComboBox cmbformtype 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   2985
         End
         Begin FLWCtrls.FWNumericTextBox fwnumericNo 
            Height          =   495
            Left            =   2160
            TabIndex        =   5
            Top             =   1320
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   873
            Max             =   1000
            Min             =   1
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   " ⁄œ«œ œ›⁄«  ç«Å"
            BeginProperty Font 
               Name            =   "B Homa"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3480
            TabIndex        =   7
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ ›—„ ·Ì»·"
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
            Left            =   3360
            TabIndex        =   6
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.CommandButton CmdDefaultSetting 
      BackColor       =   &H000000FF&
      Caption         =   " ‰ŸÌ„«  ÅÌ‘ ›—÷"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   1
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton CmdSaveSetting 
      BackColor       =   &H000000C0&
      Caption         =   "À»   ‰ŸÌ„« "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   0
      Top             =   6720
      Width           =   1935
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmbarcode.frx":A4C2
      TabIndex        =   40
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmbarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim filetemp As New FileSystemObject
Dim tempstring As TextStream
Dim Str As String
Dim LenghStr As Integer
Dim IsFileExist As Boolean
Dim LableSettingFile As String
Dim StoreNameDefine As String
Dim GoodName As String
Dim FeeUnit As String
Dim BarcodeNo As String
Dim barcodeNoView As String
Dim ColumnsPerRow As Integer
Dim XScale As Long
Dim YScale As Long
Dim Hgap As Long
Dim formtype As Integer
Dim BarcodeHeight As Long
Dim DistanceValue As Long
Dim storenamefont As String
Dim pricefont As String
Dim goodfont As String
Dim barcodefont As String

Private Declare Sub OpenPort Lib "c:\tsclib.dll" Alias _
"openport" (ByVal Command1 As String)
Private Declare Sub closeport Lib "c:\tsclib.dll" ()
Private Declare Sub sendcommand Lib "c:\tsclib.dll" ( _
ByVal command As String)
Private Declare Sub setup Lib "c:\tsclib.dll" ( _
ByVal LabelWidth As String, _
ByVal LabelHeight As String, _
ByVal Speed As String, _
ByVal Density As String, _
ByVal Sensor As String, _
ByVal Vertical As String, _
ByVal Offset As String)
Private Declare Sub downloadpcx Lib "c:\tsclib.dll" ( _
ByVal Filename As String, _
ByVal ImageName As String)
Private Declare Sub barcode Lib "c:\tsclib.dll" ( _
ByVal X As String, _
ByVal Y As String, _
ByVal CodeType As String, _
ByVal Height As String, _
ByVal Readable As String, _
ByVal rotation As String, _
ByVal Narrow As String, _
ByVal Wide As String, _
ByVal Code As String)
Private Declare Sub printerfont Lib "c:\tsclib.dll" ( _
ByVal X As String, _
ByVal Y As String, _
ByVal FontName As String, _
ByVal rotation As String, _
ByVal Xmul As String, _
ByVal Ymul As String, _
ByVal Content As String)
Private Declare Sub clearbuffer Lib "c:\tsclib.dll" ()
Private Declare Sub printlabel Lib "c:\tsclib.dll" ( _
ByVal NumberOfSet As String, _
ByVal NumberOfCopy As String)
Private Declare Sub formfeed Lib "c:\tsclib.dll" ()
Private Declare Sub nobackfeed Lib "c:\tsclib.dll" ()
Private Declare Sub windowsfont Lib "c:\tsclib.dll" ( _
ByVal X As Integer, _
ByVal Y As Integer, _
ByVal fontheight As Integer, _
ByVal rotation As Integer, _
ByVal fontstyle As Integer, _
ByVal fontunderline As Integer, _
ByVal FaceName As String, _
ByVal TextContent As String)


Private Sub CmdDefaultSetting_Click()
    FWNumericbarcode.Value = 30
    FWNumericgood.Value = 25
    FWNumericPrice.Value = 30
    FWNumericstore.Value = 30
    FWNumericSpeed.Value = 5
    FWNumericDensity.Value = 12
    fwnumericNo.Value = 1
    TxtBarcodeHeight.Text = 6
    TxtDistance.Text = 6
    TxtHeight.Text = 18
    TxtWidth.Text = 30
    TxtHGap.Text = 2
    
End Sub

Private Sub cmdprint_Click()
        
         
    If CmbPrinter.ListIndex > -1 Then
       formtype = cmbformtype.ListIndex
    Else
       MsgBox " Ìﬂ Å—Ì‰ — «‰ Œ«» ﬂ‰Ìœ"
       Exit Sub
    End If
    If cmbformtype.ListIndex = 0 Then
       MsgBox " ›—„ ·Ì»· „Ê—œ ‰Ÿ— —« «‰ Œ«» ﬂ‰Ìœ"
       Exit Sub
    End If
    
    formtype = cmbformtype.ListIndex
    
    Set tempstring = filetemp.OpenTextFile(LableFile, ForReading, False, TristateFalse)
   
   Do While tempstring.AtEndOfLine = False
     
    
        For i = 1 To 5
          
           Str = tempstring.ReadLine
           LenghStr = InStr(1, Str, "-", vbTextCompare)
           
           If InStr(1, Str, "1-", vbTextCompare) Then
              StoreNameDefine = Mid(Str, LenghStr + 1)
           
           ElseIf InStr(1, Str, "2-", vbTextCompare) Then
              GoodName = Mid(Str, LenghStr + 1)
           
           ElseIf InStr(1, Str, "3-", vbTextCompare) Then
              FeeUnit = Mid(Str, LenghStr + 1)
           
           ElseIf InStr(1, Str, "4-", vbTextCompare) Then
             barcodeNoView = Mid(Str, LenghStr + 1)
           
           ElseIf InStr(1, Str, "5-", vbTextCompare) Then
             BarcodeNo = Mid(Str, LenghStr + 1)
           
           End If
        Next i
    
        Call Doprint

    Loop

    tempstring.Close
End Sub
Sub Doprint()


    storenamefont = FWNumericstore.Value
    goodfont = FWNumericgood.Value
    barcodefont = FWNumericbarcode.Value
    pricefont = FWNumericPrice.Value
    
    XScale = Val(TxtWidth.Text) * 8
    YScale = Val(TxtHeight.Text) * 8
    Hgap = Val(TxtHGap.Text) * 8
    BarcodeHeight = Val(TxtBarcodeHeight.Text) * 8
    DistanceValue = Val(TxtDistance.Text) * 8


        
    Call OpenPort(CmbPrinter.Text)
    '    Call setup("95", "18", "3", "10", "0", "3", "3")
    '''    Call formfeed
    Call setup(((XScale + Hgap) / 8) * (formtype + 1), TxtHeight.Text, FWNumericSpeed.Value, FWNumericDensity.Value, "0", "2", "0")
    Call clearbuffer
    If StoreNameDefine = "" Then
        storenamefont = 0
    Else
        Call windowsfont(DistanceValue, 5, storenamefont, 0, 2, 0, "bmitra", Left(StoreNameDefine, 20))
        Call windowsfont((XScale + Hgap) + DistanceValue, 5, storenamefont, 0, 2, 0, "bmitra", Left(StoreNameDefine, 20))
        If formtype = 2 Then
            Call windowsfont((XScale + Hgap) * 2 + DistanceValue, 5, storenamefont, 0, 2, 0, "bmitra", Left(StoreNameDefine, 20))
        End If
    End If
    If GoodName = "" Then
        goodfont = 0
    Else
        Call windowsfont(DistanceValue, 5 + storenamefont, goodfont, 0, 2, 0, "bmitra", Left(GoodName, 20))
        Call windowsfont((XScale + Hgap) + DistanceValue, 5 + storenamefont, goodfont, 0, 2, 0, "bmitra", Left(GoodName, 20))
        If formtype = 2 Then
            Call windowsfont((XScale + Hgap) * 2 + DistanceValue, 5 + storenamefont, goodfont, 0, 2, 0, "bmitra", Left(GoodName, 20))
        End If
    End If
    If barcodeNoView = "" Then
        barcodefont = 0
    Else
        Call windowsfont(DistanceValue, BarcodeHeight + 5 + storenamefont + goodfont, barcodefont, 0, 2, 0, "bmitra", barcodeNoView)
        Call windowsfont((XScale + Hgap) + DistanceValue, BarcodeHeight + 5 + storenamefont + goodfont, barcodefont, 0, 2, 0, "bmitra", barcodeNoView)
        If formtype = 2 Then
            Call windowsfont((XScale + Hgap) * 2 + DistanceValue, BarcodeHeight + 5 + storenamefont + goodfont, barcodefont, 0, 2, 0, "bmitra", barcodeNoView)
        End If
    End If
    
    If FeeUnit = "" Then
        pricefont = 0
    Else
        Call windowsfont(DistanceValue, BarcodeHeight + 5 + storenamefont + goodfont + barcodefont, pricefont, 0, 2, 0, "bmitra", FeeUnit)
        Call windowsfont((XScale + Hgap) + DistanceValue, BarcodeHeight + 5 + storenamefont + goodfont + barcodefont, pricefont, 0, 2, 0, "bmitra", FeeUnit)
        If formtype = 2 Then
            Call windowsfont((XScale + Hgap) * 2 + DistanceValue, BarcodeHeight + 5 + storenamefont + goodfont + barcodefont, pricefont, 0, 2, 0, "bmitra", FeeUnit)
        End If
    End If
    
    
    If Len(BarcodeNo) <= 6 Then
                       
        Call barcode(DistanceValue, 5 + storenamefont + goodfont, "39", BarcodeHeight, "0", "0", "2", "3", BarcodeNo)
        Call barcode((XScale + Hgap) + DistanceValue, 5 + storenamefont + goodfont, "39", BarcodeHeight, "0", "0", "2", "3", BarcodeNo)
        If formtype = 2 Then
            Call barcode((XScale + Hgap) * 2 + DistanceValue, 5 + storenamefont + goodfont, "39", BarcodeHeight, "0", "0", "2", "3", BarcodeNo)
        End If
    Else

        Call barcode(DistanceValue, 5 + storenamefont + goodfont, "39", BarcodeHeight, "0", "0", "1", "2", BarcodeNo)
        Call barcode((XScale + Hgap) + DistanceValue, 5 + storenamefont + goodfont, "39", BarcodeHeight, "0", "0", "1", "2", BarcodeNo)
        If formtype = 2 Then
            Call barcode((XScale + Hgap) * 2 + DistanceValue, 5 + storenamefont + goodfont, "39", BarcodeHeight, "0", "0", "1", "2", BarcodeNo)
        End If
    End If
    
    Call printlabel("1", fwnumericNo.Value)
    Call closeport
   

End Sub


Private Sub CmdSaveSetting_Click()
    IsFileExist = filetemp.FileExists(LableSettingFile)
            
    If IsFileExist = False Then
       filetemp.CreateTextFile LableSettingFile
    End If
      
    Set tempstring = filetemp.OpenTextFile(LableSettingFile, ForWriting, False, TristateFalse)
    tempstring.WriteLine ("NumericNo =" & fwnumericNo.Value)
    tempstring.WriteLine ("NumericStore =" & FWNumericstore.Value)
    tempstring.WriteLine ("NumericGood =" & FWNumericgood.Value)
    tempstring.WriteLine ("NumericBarcode =" & FWNumericbarcode.Value)
    tempstring.WriteLine ("NumericPrice =" & FWNumericPrice.Value)
    tempstring.WriteLine ("NumericSpeed =" & FWNumericSpeed.Value)
    tempstring.WriteLine ("NumericDensity =" & FWNumericDensity.Value)
    tempstring.WriteLine ("TextDistance =" & TxtDistance.Text)
    tempstring.WriteLine ("TextHeight =" & TxtHeight.Text)
    tempstring.WriteLine ("TextWidth =" & TxtWidth.Text)
    tempstring.WriteLine ("TextHGap =" & TxtHGap.Text)
    tempstring.WriteLine ("TextbarcodeHeight =" & TxtBarcodeHeight.Text)
    
    tempstring.Close

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set cnn = Nothing

    SaveSetting strMainKey, "frmbarcode", "Left", Me.Left
    SaveSetting strMainKey, "frmbarcode", "Top", Me.Top


End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub LoadDefaultSetting()
    Dim Str As String
    Dim LenghStr As Integer
    Dim IsFileExist As Boolean
     
    IsFileExist = filetemp.FileExists(LableSettingFile)
    
    If IsFileExist = False Then
       CmdDefaultSetting_Click
       CmdSaveSetting_Click
    
    Else
        Set tempstring = filetemp.OpenTextFile(LableSettingFile, ForReading, False, TristateFalse)
        
        Do While tempstring.AtEndOfLine = False
           Str = tempstring.ReadLine
           LenghStr = InStr(1, Str, "=", vbTextCompare)
           
           If InStr(1, Str, "NumericNo", vbTextCompare) Then
              fwnumericNo.Value = Mid(Str, LenghStr + 1)
           ElseIf InStr(1, Str, "NumericStore", vbTextCompare) Then
              FWNumericstore.Value = Mid(Str, LenghStr + 1)
           ElseIf InStr(1, Str, "NumericGood", vbTextCompare) Then
              FWNumericgood.Value = Mid(Str, LenghStr + 1)
           ElseIf InStr(1, Str, "NumericBarcode", vbTextCompare) Then
              FWNumericbarcode.Value = Mid(Str, LenghStr + 1)
           ElseIf InStr(1, Str, "NumericPrice", vbTextCompare) Then
              FWNumericPrice.Value = Mid(Str, LenghStr + 1)
           ElseIf InStr(1, Str, "NumericSpeed", vbTextCompare) Then
              FWNumericSpeed.Value = Mid(Str, LenghStr + 1)
           ElseIf InStr(1, Str, "NumericDensity", vbTextCompare) Then
              FWNumericDensity.Value = Mid(Str, LenghStr + 1)
           ElseIf InStr(1, Str, "TextDistance", vbTextCompare) Then
              TxtDistance.Text = CStr(Mid(Str, LenghStr + 1))
           ElseIf InStr(1, Str, "TextHeight", vbTextCompare) Then
              TxtHeight.Text = CStr(Mid(Str, LenghStr + 1))
           ElseIf InStr(1, Str, "TextWidth", vbTextCompare) Then
              TxtWidth.Text = CStr(Mid(Str, LenghStr + 1))
           ElseIf InStr(1, Str, "TextHGap", vbTextCompare) Then
              TxtHGap.Text = CStr(Mid(Str, LenghStr + 1))
           ElseIf InStr(1, Str, "TextbarcodeHeight", vbTextCompare) Then
              TxtBarcodeHeight.Text = CStr(Mid(Str, LenghStr + 1))
           
           
           End If
        Loop
        tempstring.Close
    End If
End Sub


Private Sub Form_Load()
    
    CenterCenter Me
    LableSettingFile = App.Path & "\LableSettingFile.txt"
  
    cmbformtype.AddItem " "
    cmbformtype.AddItem "2 ⁄œœÌ"
    cmbformtype.AddItem "3 ⁄œœÌ"
    cmbformtype.ListIndex = 0
    

    Dim X, Printer
    For Each X In Printers
       CmbPrinter.AddItem X.DeviceName
      CmbPrinter.ListIndex = 0
    Next
    LoadDefaultSetting
    formloadFlag = False
    Me.Left = Val(GetSetting(strMainKey, "frmbarcode", "Left"))
    If Val(GetSetting(strMainKey, "frmbarcode", "Height")) > 0 Then
        Me.Height = Val(GetSetting(strMainKey, "frmbarcode", "Height"))
    End If
    If Val(GetSetting(strMainKey, "frmbarcode", "Width")) > 0 Then
        Me.Width = Val(GetSetting(strMainKey, "frmbarcode", "Width"))
    End If
    Me.Top = Val(GetSetting(strMainKey, "frmbarcode", "Top"))
    formloadFlag = True


End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub



Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, "frmbarcode", "Height", Me.Height
        SaveSetting strMainKey, "frmbarcode", "Width", Me.Width
    End If


End Sub
