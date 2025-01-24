VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmAutodiscount 
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8895
   Icon            =   "frmAutoDiscount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8895
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000000C0&
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
      Height          =   615
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CheckBox checkActive 
      Caption         =   "›⁄«·"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text1 
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
      Height          =   495
      Index           =   3
      Left            =   5880
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   3
      Left            =   3600
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text3 
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
      ForeColor       =   &H000000C0&
      Height          =   495
      Index           =   3
      Left            =   1200
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      Height          =   495
      Index           =   2
      Left            =   5880
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   2
      Left            =   3600
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text3 
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
      ForeColor       =   &H000000C0&
      Height          =   495
      Index           =   2
      Left            =   1200
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      Height          =   495
      Index           =   1
      Left            =   5880
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   1
      Left            =   3600
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text3 
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
      ForeColor       =   &H000000C0&
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   2640
      Width           =   855
   End
   Begin VB.ComboBox comboDiscount 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "frmAutoDiscount.frx":A4C2
      Left            =   3600
      List            =   "frmAutoDiscount.frx":A4C4
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   1665
   End
   Begin VB.TextBox Text3 
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
      ForeColor       =   &H000000C0&
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Index           =   0
      Left            =   3600
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text1 
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
      Height          =   495
      Index           =   0
      Left            =   5880
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1920
      Width           =   855
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   7560
      Top             =   0
      Width           =   1380
      _ExtentX        =   2434
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
      Left            =   0
      OleObjectBlob   =   "frmAutoDiscount.frx":A4C6
      TabIndex        =   33
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   32
      Top             =   2720
      Width           =   255
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   31
      Top             =   3580
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   30
      Top             =   4300
      Width           =   255
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   29
      Top             =   1880
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   ": ê—œ ‘Êœ »Â "
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   2160
      TabIndex        =   27
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   ": ê—œ ‘Êœ »Â "
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   2160
      TabIndex        =   26
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   ": ê—œ ‘Êœ »Â "
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   2160
      TabIndex        =   25
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   ": ê—œ ‘Êœ »Â "
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   2160
      TabIndex        =   24
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblTitel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ‰ŸÌ„  Œ›Ì› « Ê„« Ìﬂ"
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
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "«“ „»·€"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   22
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   " « „»·€"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   21
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "«“ „»·€"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   20
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " « „»·€"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   19
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "«“ „»·€"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   18
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " « „»·€"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "«—ﬁ«„ ê—œ ò—œ‰"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblto 
      BackStyle       =   0  'Transparent
      Caption         =   "   « „»·€ "
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblfrom 
      BackStyle       =   0  'Transparent
      Caption         =   "«“ „»·€"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "frmAutodiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim cnn As New ADODB.Connection
Dim Parameter() As Parameter

Private Sub CancelButton_Click()
    ExitForm
End Sub

Private Sub comboDiscount_Click()
LoadDataAutoDiscount
End Sub

Private Sub Form_Load()
   MyFormAddEditMode = ViewMode
   SetFirstToolBar
   If cnn.State = adStateClosed Then cnn.ConnectionString = strConnectionString
   
   If cnn.State <> adStateOpen Then cnn.Open
    
    comboDiscount.Clear
    comboDiscount.AddItem "1"
    comboDiscount.ItemData(comboDiscount.NewIndex) = 0
    comboDiscount.AddItem "2"
    comboDiscount.ItemData(comboDiscount.NewIndex) = 1
    comboDiscount.AddItem "3"
    comboDiscount.ItemData(comboDiscount.NewIndex) = 2
    comboDiscount.AddItem "4"
    comboDiscount.ItemData(comboDiscount.NewIndex) = 3
    
    If comboDiscount.ListCount > 0 Then comboDiscount.ListIndex = 0

    LoadForm Me.Name
    CenterCenterOffset Me

    
    End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Rc = Nothing
    Set rctmp = Nothing
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
        
    VarActForm = ""
    Unload frmAutodiscount
    
    mdifrm.Toolbar1.Buttons(20).Enabled = False
    mdifrm.Toolbar1.Buttons(21).Enabled = False
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
    LoadDataAutoDiscount
    
End Sub

Public Sub ExitForm()
    Unload Me
End Sub

Public Sub LoadDataAutoDiscount()
    Dim Rst As New ADODB.Recordset
    Dim i As Integer
   i = 0
''    Select Case comboDiscount.ListIndex
''
''        Case 0
'
            ReDim Parameters(0) As Parameter
            
            Parameters(0) = GenerateInputParameter("@Type", adInteger, 4, comboDiscount.ListIndex)
            Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_AutoDiscount", Parameters)
                If Not (Rst.EOF = True And Rst.BOF = True) Then
                    While Rst.EOF <> True

                   Text1(i).Text = Rst!FromNumber
                   Text2(i).Text = Rst!ToNumber
                   Text3(i).Text = Rst!RoundNumber
                   i = i + 1
                      checkActive.Value = Rst!Active
                        Rst.MoveNext
                    Wend
             ''  checkActive.Value = Rst!Active
               
                End If

''        Case 1
'''
''
''        Case 2
''
''
''    End Select

    If Rst.State = 1 Then Rst.Close
End Sub

Public Sub SetFirstToolBar()
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(1).Enabled = False   'Home
    mdifrm.Toolbar1.Buttons(2).Enabled = False   'PageUp
    mdifrm.Toolbar1.Buttons(3).Enabled = False   'PageDown
    mdifrm.Toolbar1.Buttons(4).Enabled = False   'End
        
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    If MyFormAddEditMode = ViewMode Then  ' View Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = True 'Delete
     
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False 'Delete

    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False 'Delete
        
    End If

    HeaderLabel Val(MyFormAddEditMode), fwlblMode
End Sub
Public Sub Cancel()

    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    checkActive.Enabled = False
    LoadDataAutoDiscount
    
End Sub
Public Sub Edit()
    MyFormAddEditMode = EditMode
    mdifrm.Toolbar1.Buttons(7).Enabled = False
  
    SetFirstToolBar
   '' vsFlex.Editable = flexEDKbdMouse
    checkActive.Enabled = True
End Sub

Public Sub Update()
    
    ReDim Parameter(3) As Parameter
   
  
       
                    For i = 0 To 3
                   ReDim Parameter(5) As Parameter
                            Parameter(0) = GenerateInputParameter("@FromNumber", adDouble, 4, Val(Text1(i).Text))
                            Parameter(1) = GenerateInputParameter("@ToNumber", adDouble, 4, Val(Text2(i).Text))
                            Parameter(2) = GenerateInputParameter("@RoundNumber", adDouble, 50, Val(Text3(i).Text))
                            Parameter(3) = GenerateInputParameter("@RowNumber", adVarChar, 50, Str(comboDiscount.ListIndex) + LTrim(Str(i + 1)))
                            Parameter(4) = GenerateInputParameter("@Active", adInteger, 4, checkActive.Value)
                            Parameter(5) = GenerateOutputParameter("@Check", adInteger, 4)
                            Result = RunParametricStoredProcedure("Update_tblTotal_AutoDiscount", Parameter)
                       
                    Next
              
            If Result <> -1 Then
            
                frmMsg.fwlblMsg.Caption = " €ÌÌ— «ÿ·«⁄«  »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
            Else

                frmMsg.fwlblMsg.Caption = "„ «”›«‰Â «ÿ·«⁄«   €ÌÌ— ‰Ì«› . ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                
            End If
            
    

    
      mdifrm.Toolbar1.Buttons(7).Enabled = True
      MyFormAddEditMode = ViewMode
      SetFirstToolBar
      checkActive.Enabled = False
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

