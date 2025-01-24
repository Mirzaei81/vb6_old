VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTables2 
   BackColor       =   &H00C0FFC0&
   Caption         =   "                                                                                                                 áíÓÊ ãíÒåÇ"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "Nazanin"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTables2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   -120
      Width           =   11415
      Begin VB.CheckBox chkViewSpecial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "ÝíáÊÑ ÈÑ ÇÓÇÓ äÝÑ"
         Height          =   420
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
      Begin FLWCtrls.FWNumericTextBox txtInterval 
         Height          =   480
         Left            =   840
         TabIndex        =   23
         Top             =   200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   847
         Min             =   1
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox TxtFontSize 
         Alignment       =   1  'Right Justify
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
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Text            =   "12"
         Top             =   240
         Width           =   720
      End
      Begin VB.TextBox TxtHeight 
         Alignment       =   1  'Right Justify
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
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Text            =   "850"
         Top             =   225
         Width           =   720
      End
      Begin VB.TextBox TxtWidth 
         Alignment       =   1  'Right Justify
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
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Text            =   "1400"
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÓÇíÒ ÝæäÊ"
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
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblIntervalTitle 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÒãÇä ÈÑæÒ ÑÓÇäí :"
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
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblSecondTitle 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ËÇäíå"
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
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÚÑÖ"
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
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Øæá"
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
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   0
   End
   Begin VB.Frame frameMap 
      Height          =   600
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   7920
      Width           =   11415
      Begin VB.PictureBox Picture4 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4800
         RightToLeft     =   -1  'True
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000000C0&
         Height          =   255
         Left            =   2160
         RightToLeft     =   -1  'True
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000080FF&
         Height          =   255
         Left            =   7560
         RightToLeft     =   -1  'True
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H0000C000&
         Height          =   255
         Left            =   10320
         RightToLeft     =   -1  'True
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.Label LblOverTime 
         Alignment       =   1  'Right Justify
         Caption         =   "ãíÒ ÈÇ ÊÇíã ÇÖÇÝí"
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
         Index           =   0
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblTableWithInvoicePrint 
         Alignment       =   1  'Right Justify
         Caption         =   "Ç ÝÇßÊæÑÑÝÊå ÔÏå"
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
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblReservedTables 
         Alignment       =   1  'Right Justify
         Caption         =   "ãíÒ Ñ"
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
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblEmptyTables 
         Alignment       =   1  'Right Justify
         Caption         =   "ãíÒ ÎÇáí"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   13
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   16777088
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Ñæå Çæá"
      TabPicture(0)   =   "frmTables2.frx":A4C2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmTables2.frx":A4DE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmTables2.frx":A4FA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmTables2.frx":A516
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Tab 4"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Tab 5"
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Tab 6"
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Tab 7"
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Tab 8"
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "Tab 9"
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "Tab 10"
      Tab(10).ControlEnabled=   0   'False
      Tab(10).ControlCount=   0
      TabCaption(11)  =   "Tab 11"
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
      TabCaption(12)  =   "Tab 12"
      Tab(12).ControlEnabled=   0   'False
      Tab(12).ControlCount=   0
      Begin VB.Frame Frame1 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1080
         Width           =   11295
         Begin FLWCtrls.FWCoolButton cmd 
            Height          =   855
            Index           =   0
            Left            =   360
            TabIndex        =   2
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmTables2.frx":A532
      TabIndex        =   10
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmTables2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Parameter() As Parameter
Dim Rst As Recordset
Private Type Position
    X As Long
    Y As Long
End Type
Dim lastPosition(0 To 4) As Position
Dim i, j, k, m, l, i0, i1, i2, i3, i4 As Integer
Dim varCmd As CommandButton
Dim varlbl As Label
Dim tablesCount As Integer
Dim Part(0 To 4) As Integer

Private Sub chkViewSpecial_Click()
    GetTables CurrentBranch
End Sub



Private Sub Form_Activate()
    GetCount CurrentBranch
    If tablesCount > 0 Then
        If clsArya.HardLockSerialNo = "92072202873" Then
            chkViewSpecial.Value = 1
        Else
            GetTables CurrentBranch
        End If
'''''        DetectReservedTables CurrentBranch
        DetectBusyTables CurrentBranch
        DetectOtherTables CurrentBranch
    End If
    For i = 0 To 4
        If Part(i) = clsStation.PartitionID Then
            SSTab1.Tab = i
            Exit For
        End If
    Next
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 27  ' Esc
            
                  Unload Me
              End Select
          Case 2
               Select Case KeyCode
                  Case 123  'Exit
                     If clsStation.KeyboardType = EnumKeyBoardType.Rb2 Then
                        Unload Me
                     End If
              End Select

    End Select

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Private Sub EmptyTable(TableNo As Integer)
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intTableNo", adInteger, 4, TableNo)
    RunParametricStoredProcedure "Update_tTable_By_Empty", Parameter

End Sub
Private Sub label_Click(Index As Integer)
    If Part(0) = clsStation.PartitionID Then
        If CheckTableState(Label(Index).Tag) = True Then      'If empty
            mvarTable = Val(Label(Index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label(Index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label(Index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub

Private Sub Label1_Click(Index As Integer)
'    If Part(1) = clsStation.PartitionID Then
        If CheckTableState(Label1(Index).Tag) = True Then
            mvarTable = Val(Label1(Index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label1(Index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label1(Index).Tag, CurrentBranch
        End If
        Unload Me
'    Else
'        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
'    End If
End Sub

Private Sub label2_Click(Index As Integer)
'    If Part(2) = clsStation.PartitionID Then
        If CheckTableState(Label2(Index).Tag) = True Then
            mvarTable = Val(Label2(Index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label2(Index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label2(Index).Tag, CurrentBranch
        End If
        Unload Me
'    Else
'        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
'    End If
End Sub
Private Sub label3_Click(Index As Integer)
'    If Part(3) = clsStation.PartitionID Then
        If CheckTableState(Label3(Index).Tag) = True Then
            mvarTable = Val(Label3(Index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label3(Index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label3(Index).Tag, CurrentBranch
        End If
        Unload Me
'    Else
'        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
'    End If
End Sub
Private Sub label4_Click(Index As Integer)
'    If Part(4) = clsStation.PartitionID Then
        If CheckTableState(Label4(Index).Tag) = True Then
            mvarTable = Val(Label4(Index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label4(Index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label4(Index).Tag, CurrentBranch
        End If
        Unload Me
'    Else
'        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
'    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    formloadFlag = False
    
    If Val(GetSetting(strMainKey, Me.Name, "TimerInterval")) > 0 Then
        Timer1.Interval = Val(GetSetting(strMainKey, Me.Name, "TimerInterval"))
    Else
         Timer1.Interval = 5000
    End If
    If Val(GetSetting(strMainKey, Me.Name, "TxtWidth")) > 0 Then
        TxtWidth = Val(GetSetting(strMainKey, Me.Name, "TxtWidth"))
    Else
         TxtWidth = 1600
    End If
    If Val(GetSetting(strMainKey, Me.Name, "TxtHeight")) > 0 Then
        TxtHeight = Val(GetSetting(strMainKey, Me.Name, "TxtHeight"))
    Else
         TxtHeight = 850
    End If
    If Val(GetSetting(strMainKey, Me.Name, "TxtFontSize")) > 0 Then
        TxtFontSize = Val(GetSetting(strMainKey, Me.Name, "TxtFontSize"))
    Else
         TxtFontSize = 12
    End If
    txtInterval.Value = CStr(Timer1.Interval / 1000)
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_tPartitions", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        'Debug.Print Rst.RecordCount
        i = 0
        While Rst.EOF <> True
            i = i + 1
            Rst.MoveNext
        Wend
    End If
    
    SSTab1.Tabs = i
'    For i = 0 To SSTab1.Tabs
'        SSTab1.TabVisible(i) = False
'    Next
    
    Dim Rst As New ADODB.Recordset
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_tPartitions", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        'Debug.Print Rst.RecordCount
        i = 0
        While Rst.EOF <> True
            SSTab1.TabVisible(i) = True
            SSTab1.TabCaption(i) = Rst.Fields("PartitionDescription").Value
            Part(i) = Rst.Fields("PartitionID").Value
            i = i + 1
            Rst.MoveNext
        Wend
    End If
    

    Me.Left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 0 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 0 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.Top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    
    formloadFlag = True

Exit Sub
ErrHandler:
    formloadFlag = True
    ShowDisMessage err.Description, 2000
End Sub
Private Sub GetCount(Branch As Integer)
    On Error GoTo ErrHandler
        ReDim Parameter(0)
        Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, Branch)
        
        Set Rst = RunParametricStoredProcedure2Rec("GetTableCountByBranch", Parameter)
        
        If Not (Rst.EOF And Rst.BOF) Then
            Do While Rst.EOF = False
                tablesCount = Rst!ct
                Rst.MoveNext
            Loop
        End If
        
    Exit Sub
ErrHandler:
    MsgBox err.Description
    LogSave "frmTables", err, "GetCount"
End Sub



Private Sub GetTables(Branch As Integer)
    On Error GoTo ErrHandler
    
    For i = 1 To Label.Count - 1
        Unload Label(i)
    Next
    For i = 1 To Label1.Count - 1
        Unload Label1(i)
    Next
    For i = 1 To Label2.Count - 1
        Unload Label2(i)
    Next
    For i = 1 To Label3.Count - 1
        Unload Label3(i)
    Next
    For i = 1 To Label4.Count - 1
        Unload Label4(i)
    Next
    
    Label(0).Width = TxtWidth
    Label(0).Height = TxtHeight
    Label(0).Font.Size = TxtFontSize
'    Label(0).ForeColor = vbBlue
    Label1(0).Width = TxtWidth
    Label1(0).Height = TxtHeight
    Label1(0).Font.Size = TxtFontSize
    Label2(0).Width = TxtWidth
    Label2(0).Height = TxtHeight
    Label2(0).Font.Size = TxtFontSize
    Label3(0).Width = TxtWidth
    Label3(0).Height = TxtHeight
    Label3(0).Font.Size = TxtFontSize
    
    lastPosition(0).X = 360
    lastPosition(0).Y = 360
    lastPosition(1).X = 360
    lastPosition(1).Y = 360
    lastPosition(2).X = 360
    lastPosition(2).Y = 360
    lastPosition(3).X = 360
    lastPosition(3).Y = 360
    lastPosition(4).X = 360
    lastPosition(4).Y = 360
'    Label(0).Height = 1000
'    Label(1).Height = 1000
'    Label(2).Height = 1000
'    Label(3).Height = 1000
'    Label(4).Height = 1000
    
    ReDim Parameter(1)
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, Branch)
    Parameter(1) = GenerateInputParameter("@TableControl", adInteger, 4, 0)
    
    Set Rst = RunParametricStoredProcedure2Rec("RetriveTable_Branch", Parameter)
    i = 1
    i0 = 1
    i1 = 1
    i2 = 1
    i3 = 1
    i4 = 1
    If Not (Rst.EOF And Rst.BOF) Then
        Do While Rst.EOF = False
            If Not (chkViewSpecial.Value = 1 And Val(frmInvoice.TxtGuestNo) <> Rst!NumberOfChair) Then
                Select Case Rst!PartitionID
                    Case Part(0)
                        If lastPosition(0).X > Frame1.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(0).X = 360
                            lastPosition(0).Y = lastPosition(0).Y + Val(TxtHeight) + 105
                        End If
                        Load Label(i0)
                        Label(i0).Left = lastPosition(0).X
                        Label(i0).Top = lastPosition(0).Y
                        Label(i0).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label(i0).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label(i0).Caption = Rst!TableDescription
                        End If
                        Label(i0).Tag = Rst!No
                        Label(i0).Width = Val(TxtWidth)
                        Label(i0).Height = Val(TxtHeight)
                        Label(i0).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                           'Not empty
                        'Label(i).Enabled = False
                        lastPosition(0).X = lastPosition(0).X + Val(TxtWidth) + 105
                        i0 = i0 + 1
                    Case Part(1)
                        If lastPosition(1).X > 10500 Then
                            lastPosition(1).X = 360
                            lastPosition(1).Y = lastPosition(1).Y + Val(TxtHeight) + 105
                        End If
                        Load Label1(i1)
                        Label1(i1).Left = lastPosition(1).X
                        Label1(i1).Top = lastPosition(1).Y
                        Label1(i1).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label1(i1).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label1(i1).Caption = Rst!TableDescription
                        End If
                        Label1(i1).Tag = Rst!No
                        Label1(i1).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                   'Not empty
                        'Label(i).Enabled = False
                        lastPosition(1).X = lastPosition(1).X + Val(TxtWidth) + 105
                        i1 = i1 + 1
                    Case Part(2)
                        If lastPosition(2).X > 10500 Then
                            lastPosition(2).X = 360
                            lastPosition(2).Y = lastPosition(2).Y + Val(TxtHeight) + 105
                        End If
                        Load Label2(i2)
                        Label2(i2).Left = lastPosition(2).X
                        Label2(i2).Top = lastPosition(2).Y
                        Label2(i2).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label2(i2).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label2(i2).Caption = Rst!TableDescription
                        End If
                        Label2(i2).Tag = Rst!No
                        Label2(i2).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(2).X = lastPosition(2).X + Val(TxtWidth) + 105
                        i2 = i2 + 1
                    Case Part(3)
                        If lastPosition(3).X > 10500 Then
                            lastPosition(3).X = 360
                            lastPosition(3).Y = lastPosition(3).Y + Val(TxtHeight) + 105
                        End If
                        Load Label3(i3)
                        Label3(i3).Left = lastPosition(3).X
                        Label3(i3).Top = lastPosition(3).Y
                        Label3(i3).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label3(i3).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label3(i3).Caption = Rst!TableDescription
                        End If
                        Label3(i3).Tag = Rst!No
                        Label3(i3).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(3).X = lastPosition(3).X + Val(TxtWidth) + 105
                        i3 = i3 + 1
                     Case Part(4)
                        If lastPosition(4).X > 10500 Then
                            lastPosition(4).X = 360
                            lastPosition(4).Y = lastPosition(4).Y + Val(TxtHeight) + 105
                        End If
                        Load Label4(i4)
                        Label4(i4).Left = lastPosition(4).X
                        Label4(i4).Top = lastPosition(4).Y
                        Label4(i4).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label4(i4).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label4(i4).Caption = Rst!TableDescription
                        End If
                        Label4(i4).Tag = Rst!No
                        Label4(i4).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(4).X = lastPosition(4).X + Val(TxtWidth) + 105
                        i4 = i4 + 1
                End Select
            End If
            Rst.MoveNext
        Loop
    
    End If

    Exit Sub
ErrHandler:
   MsgBox err.Description
   LogSave "frmTables", err, "GetNumberTables"
End Sub
'Private Sub DetectReservedTables(Branch As Integer)
'    On Error GoTo ErrHandler
'    ReDim Parameter(1)
'    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, Branch)
'    Parameter(1) = GenerateInputParameter("@TableControl", adInteger, 4, 0)
'
'    Set Rst = RunParametricStoredProcedure2Rec("RetriveTable_Branch", Parameter)
'
'    If Not (Rst.EOF And Rst.BOF) Then
'            For Each varCmd In cmd    'Enable the empty tables
'                 If varCmd.Tag = CStr(Rst!No) Then
'                     varCmd.BackColor = &HA0B5FF            'red for reserved
'                     varCmd.Enabled = True
'                     Rst.MoveNext
'                     If Rst.EOF Then Exit For
'                End If
'            Next
'            If Rst.EOF = False Then
'                For Each varCmd In cmd1
'                     If varCmd.Tag = CStr(Rst!No) Then
'                         varCmd.BackColor = &HA0B5FF
'                         varCmd.Enabled = True
'                         Rst.MoveNext
'                         If Rst.EOF Then Exit For
'                    End If
'                Next
'            End If
'           If Rst.EOF = False Then
'                For Each varCmd In cmd2
'                     If varCmd.Tag = CStr(Rst!No) Then
'                         varCmd.BackColor = &HA0B5FF
'                         varCmd.Enabled = True
'                         Rst.MoveNext
'                         If Rst.EOF Then Exit For
'                    End If
'                Next
'            End If
'
'    End If
''    SSTab1.Tab = 0
'    Exit Sub
'ErrHandler:
'    MsgBox err.Description
'    LogSave "frmTables", err, "DetectEmptyTables"
'End Sub

Private Sub DetectEmptyTables(Branch As Integer)
    On Error GoTo ErrHandler
    ReDim Parameter(1)
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, Branch)
    Parameter(1) = GenerateInputParameter("@TableControl", adInteger, 4, 0)
    
    Set Rst = RunParametricStoredProcedure2Rec("RetriveTable_Branch", Parameter)

    If Not (Rst.EOF And Rst.BOF) Then
            For Each varlbl In Label    'Enable the empty tables
                If varlbl.Tag = CStr(Rst!No) Then
                    If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                    Rst.MoveNext
                    If Rst.EOF Then Exit For
                End If
            Next
            If Rst.EOF = False Then
                For Each varlbl In Label1
                    If varlbl.Tag = CStr(Rst!No) Then
                        If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                        Rst.MoveNext
                        If Rst.EOF Then Exit For
                    End If
                Next
            End If
           If Rst.EOF = False Then
                For Each varlbl In Label2
                    If varlbl.Tag = CStr(Rst!No) Then
                        If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                        Rst.MoveNext
                        If Rst.EOF Then Exit For
'                    Else
'                        varlbl.BackColor = &HA0B5FF
'                        varlbl.Enabled = True
                    End If
                Next
            End If
           If Rst.EOF = False Then
                For Each varlbl In Label3
                    If varlbl.Tag = CStr(Rst!No) Then
                        If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                        Rst.MoveNext
                        If Rst.EOF Then Exit For
                    End If
                Next
            End If
           If Rst.EOF = False Then
                For Each varlbl In Label4
                    If varlbl.Tag = CStr(Rst!No) Then
                        If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                        Rst.MoveNext
                        If Rst.EOF Then Exit For
                    End If
                Next
            End If

    End If
'    SSTab1.Tab = 0
    Exit Sub
ErrHandler:
    MsgBox err.Description
    LogSave "frmTables", err, "DetectResevedTables"
End Sub

Private Sub DetectBusyTables(Branch As Integer)
    On Error GoTo ErrHandler
    ReDim Parameter(0)
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, Branch)

    Set Rst = RunParametricStoredProcedure2Rec("Get_tblSamar_TableUsage_BusyTable", Parameter)

    If Not (Rst.EOF And Rst.BOF) Then
            For Each varlbl In Label    'Enable the empty tables
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Val(Rst!MinuteUseDiff) >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                        
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
'                Else
'                    varlbl.BackColor = &HA0B5FF         'red for reserved
'                    varlbl.Enabled = True
                End If
            Next
            If Rst.EOF = False Then
                For Each varlbl In Label1
                     If varlbl.Tag = CStr(Rst!No) Then
                        varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"             'green for empty
                        If Not IsNull(Rst!nvcMaxUseTime) Then
                           If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                        End If
                        Rst.MoveNext
                        If Rst.EOF Then Exit For
'                    Else
'                        varlbl.BackColor = &HA0B5FF
'                        varlbl.Enabled = True
                    End If
                Next
            End If
           If Rst.EOF = False Then
                For Each varlbl In Label2
                     If varlbl.Tag = CStr(Rst!No) Then
                        varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                        If Not IsNull(Rst!nvcMaxUseTime) Then
                           If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                        End If
                        Rst.MoveNext
                        If Rst.EOF Then Exit For
'                    Else
'                        varlbl.BackColor = &HA0B5FF
'                        varlbl.Enabled = True
                    End If
                Next
            End If
           If Rst.EOF = False Then
                For Each varlbl In Label3
                    If varlbl.Tag = CStr(Rst!No) Then
                        varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"           'green for empty
                        If Not IsNull(Rst!nvcMaxUseTime) Then
                           If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                        End If
                        Rst.MoveNext
                        If Rst.EOF Then Exit For
'                    Else
'                        varlbl.BackColor = &HA0B5FF
'                        varlbl.Enabled = True
                    End If
                Next
            End If
           If Rst.EOF = False Then
                For Each varlbl In Label4
                    If varlbl.Tag = CStr(Rst!No) Then
                        varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                        If Not IsNull(Rst!nvcMaxUseTime) Then
                           If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                        End If
                        Rst.MoveNext
                        If Rst.EOF Then Exit For
'                    Else
'                        varlbl.BackColor = &HA0B5FF
'                        varlbl.Enabled = True
                    End If
                Next
            End If

    End If
'    SSTab1.Tab = 0
    Exit Sub
ErrHandler:
    MsgBox err.Description
    LogSave "frmTables", err, "DetectResevedTables"
    Resume Next
End Sub

Private Sub DetectOtherTables(Branch As Integer)
    On Error GoTo ErrHandler
        ReDim Parameter(0)
        Parameter(0) = GenerateInputParameter("@branch", adInteger, 4, Branch)
            
        Set Rst = RunParametricStoredProcedure2Rec("GetTablesWithInvoicePrint", Parameter)
        
        If Not (Rst.EOF And Rst.BOF) Then
            Do While Rst.EOF = False
                For Each varlbl In Label    'Enable the empty tables
                     If varlbl.Tag = CStr(Rst!No) Then
                         varlbl.BackColor = vbRed
                         Rst.MoveNext
                         If Rst.EOF Then Exit For
                    End If
                Next
                If Rst.EOF = False Then
                For Each varlbl In Label1
                     If varlbl.Tag = CStr(Rst!No) Then
                         varlbl.BackColor = vbRed
                         varlbl.Enabled = True
                         Rst.MoveNext
                         If Rst.EOF Then Exit For
                    End If
                Next
                End If
                If Rst.EOF = False Then
                     For Each varlbl In Label2
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label3
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label4
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                 If Rst.EOF = False Then Rst.MoveNext
            Loop
        End If
    
    Exit Sub
ErrHandler:
    MsgBox err.Description
    LogSave "frmTables", err, "DetectOtherTables"
End Sub


Private Sub FindInvoice(TableNo As Integer, Branch As Integer)
    On Error GoTo ErrHandler
    ReDim Parameter(1)
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, Branch)
    Parameter(1) = GenerateInputParameter("@TableNO", adInteger, 4, TableNo)
    
    Set Rst = RunParametricStoredProcedure2Rec("GetInvoiceByTable", Parameter)
    mvarInvoiceNO = 0
    If Not (Rst.EOF And Rst.BOF) Then
        mvarInvoiceNO = Rst!No
    End If
    Exit Sub
ErrHandler:
    MsgBox err.Description
    LogSave "frmTables", err, "FindInvoice"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting strMainKey, Me.Name, "TimerInterval", CStr(Val(txtInterval.Value) * 1000)
    SaveSetting strMainKey, Me.Name, "TxtWidth", CStr(Val(TxtWidth.Text))
    SaveSetting strMainKey, Me.Name, "TxtHeight", CStr(Val(TxtHeight.Text))
    SaveSetting strMainKey, Me.Name, "TxtFontSize", CStr(Val(TxtFontSize.Text))

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub



Private Sub Timer1_Timer()
    Timer1.Interval = txtInterval.Value * 1000
    If tablesCount > 0 Then
        DetectEmptyTables CurrentBranch
        DetectBusyTables CurrentBranch
        DetectOtherTables CurrentBranch
    End If
End Sub


'Private Sub FindTable(strTableName As String)
'    Dim Result As Boolean
'    Result = False
'    On Error GoTo ErrHandler
'        For Each varlbl In Label
'            If varlbl.Caption = strTableName Then
'                If varlbl.Enabled Then
'                    varlbl.SetFocus
'                End If
''                varlbl.BackColor = vbBlue
'                SSTab1.Tab = 0
'                Result = True
'                Exit For
'            End If
'        Next
'        If Result = False Then
'            For Each varlbl In Label1
'                If varlbl.Caption = strTableName Then
'                    If varlbl.Enabled Then varlbl.SetFocus
''                    varlbl.BackColor = vbBlue
'                    SSTab1.Tab = 1
'                    Result = True
'                    Exit For
'                End If
'            Next
'        End If
'        If Result = False Then
'            For Each varlbl In Label2
'                If varlbl.Caption = strTableName Then
'                    If varlbl.Enabled Then varlbl.SetFocus
''                    varlbl.BackColor = vbBlue
'                    SSTab1.Tab = 2
'                    Result = True
'                    Exit For
'                End If
'            Next
'        End If
'        If Result = False Then
'            For Each varlbl In Label3
'                If varlbl.Caption = strTableName Then
'                    If varlbl.Enabled Then varlbl.SetFocus
''                    varlbl.BackColor = vbBlue
'                    SSTab1.Tab = 3
'                    Result = True
'                    Exit For
'                End If
'            Next
'        End If
'        If Result = False Then
'            For Each varlbl In Label4
'                If varlbl.Caption = strTableName Then
'                    If varlbl.Enabled Then varlbl.SetFocus
''                    varlbl.BackColor = vbBlue
'                    SSTab1.Tab = 4
'                    Result = True
'                    Exit For
'                End If
'            Next
'        End If
'    Exit Sub
'ErrHandler:
'    MsgBox err.Description
'    LogSave "frmTables", err, "FindTable"
'End Sub
Private Function CheckTableState(TableNo As String) As Boolean
    On Error GoTo ErrHandler
    CheckTableState = True
        ReDim Parameter(1)
        Parameter(0) = GenerateInputParameter("@TableNO", adInteger, 4, TableNo)
        Parameter(1) = GenerateInputParameter("@nvcDate", adWChar, 8, mvarDate)
        
        Set Rst = RunParametricStoredProcedure2Rec("CheckTableStatus", Parameter)
        If Not (Rst.EOF And Rst.BOF) Then  ' Not Empty
            CheckTableState = False 'IIf(Rst!Empty = False, False, True)
        End If
    Exit Function
ErrHandler:
    MsgBox err.Description
    LogSave "FrmTables", err, "CheckTableState"
End Function


