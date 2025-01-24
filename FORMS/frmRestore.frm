VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmRestore 
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   Icon            =   "frmRestore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   10095
   Begin VB.Frame Frame7 
      Caption         =   "„”Ì— Log  ›«Ì· "
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   5880
      Width           =   4260
      Begin VB.TextBox TxtLogPath 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   4065
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "„”Ì— ›«Ì· œÌ «"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   4800
      Width           =   4260
      Begin VB.TextBox TxtDataPath 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   4065
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "‰«„ ”—Ê—"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   5880
      Width           =   2700
      Begin VB.TextBox TxtServer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Text            =   "."
         Top             =   360
         Width           =   2385
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "«‰ Œ«» »«‰ﬂ »—«Ì «Ã—«Ì ‰—„ «›“«— "
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   3975
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   720
      Width           =   4215
      Begin VB.CommandButton CmdPreSet 
         Caption         =   " ‰ŸÌ„ „ﬁ«œÌ— »Â ÅÌ‘ ›—÷"
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
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   3360
         Width           =   3375
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1560
         Width           =   735
      End
      Begin VB.ComboBox cmbDataBase 
         BackColor       =   &H8000000F&
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
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   960
         Width           =   2235
      End
      Begin VB.TextBox TxtDataSource 
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
         Height          =   420
         Left            =   120
         TabIndex        =   19
         Text            =   "."
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtPassword 
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
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   14
         Text            =   "lemon7430"
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox TxtUserName 
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
         Height          =   420
         Left            =   720
         TabIndex        =   13
         Text            =   "sa"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "‰«„ ”—Ê—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "»«‰ﬂ Â«Ì «ÿ·«⁄« Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ò«—»— œÌ «»Ì”"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   2160
         Width           =   1635
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ﬂ·„Â ⁄»Ê—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2760
         Width           =   1635
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "‰«„ œÌ « »Ì”"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   5880
      Width           =   2700
      Begin VB.TextBox TxtDbName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Text            =   "Temp"
         Top             =   360
         Width           =   2385
      End
   End
   Begin VB.CommandButton FwbtnEsc 
      BackColor       =   &H00000080&
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
      Left            =   120
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton FwbtnOk 
      BackColor       =   &H00008000&
      Caption         =   " Ê·Ìœ »«‰ﬂ ÃœÌœ"
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
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   " „”Ì— ›«Ì·Â«Ì Å‘ Ì»«‰"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   4125
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   5490
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3450
         Left            =   3165
         Pattern         =   "*.fbk;*.dbk"
         TabIndex        =   5
         Top             =   405
         Width           =   2205
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   120
         TabIndex        =   1
         Top             =   1065
         Width           =   2925
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   0
         Top             =   405
         Width           =   2925
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "›«Ì· Å‘ Ì»«‰"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4800
      Width           =   5460
      Begin VB.TextBox txtFile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   5025
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmRestore.frx":A4C2
      TabIndex        =   11
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   " ⁄—Ì› Ê »«“ê—œ«‰Ì »«‰ﬂÂ«Ì «ÿ·«⁄« Ì    "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iHeight As Integer
Private iWidth As Integer
'rivate fls As New FileSystemObject
Private cmd As New ADODB.command
'rivate mycls As FileClass
'rivate ClsGl As New ClsGl
Private Rc As New ADODB.Recordset
Private rctmp As New ADODB.Recordset
Private DbName, RestoreDataBaseData, RestoreDataBaseLog As String
Private Conn As New ADODB.Connection
Dim mvarProvider As String
Dim DataSource As String
Dim AryaSettingFile As String

Private Sub cmbDataBase_Change()
    cmbDataBase_Click
End Sub

Private Sub cmbDataBase_Click()
    If formloadFlag = False Then Exit Sub
    DbName = cmbDataBase.Text
    CmdPreSet.Enabled = False
    strConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & clsArya.DBLogin & "; Password = " & SqlPass & " ; Initial Catalog = " & DbName & ";Data Source=" & DataSource
    frmNewLoginInfo.Show vbModal
    If LoginSucceeded = False Then
        frmRestore.cmbDataBase.Clear
    Else
        Call ODBCSetting(DataSource, DbName)
        clsArya.DbName = DbName
        CmdPreSet.Enabled = True
    End If
End Sub

Private Sub FillDataBase()
    Dim i As Integer
    i = 0
    cmbDataBase.Clear
    Set rctmp = RunStoredProcedure2RecordSet("master.dbo.sp_MShasdbaccess", Conn)
    Do While rctmp.EOF = False
        cmbDataBase.AddItem rctmp!DbName
        cmbDataBase.ItemData(cmbDataBase.NewIndex) = i
        i = i + 1
        rctmp.MoveNext
    Loop
    rctmp.Close
    For i = 0 To cmbDataBase.ListCount - 1
        cmbDataBase.ListIndex = i
        If LCase(clsArya.DbName) = LCase(cmbDataBase.Text) Then
            Exit For
        End If
    Next

End Sub


Private Sub cmbDataBase_DropDown()
   'FillDataBase
End Sub

Private Sub CmdPreSet_Click()
    clsArya.ServerName = DataSource
    clsArya.DbName = DbName
    Call setAryaSettingFile
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo ErrHandler
    cmbDataBase.Clear
    DataSource = TxtDataSource.Text
    If Conn.State = 1 Then Conn.Close
    strConnectionString = ""
    Conn.Open "provider =  " + mvarProvider + "; data source = " + DataSource + ";Initial Catalog =msdb;user id=" + TxtUserName.Text + ";Password=" + txtPassword.Text
    formloadFlag = False
    FillDataBase
    formloadFlag = True

Exit Sub
ErrHandler:
    ShowMessage err.Description, True, False, " «ÌÌœ", " "
    
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    File1.Pattern = "*.*"
End Sub

Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive
 
End Sub

Private Sub File1_Click()
    If Len(File1.Path) <> 3 Then
        txtFile.Text = File1.Path & "\" & File1.Filename
    Else
        txtFile.Text = File1.Path & File1.Filename
    End If
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Load()
Dim shamsi As String
 
    CenterTop Me

    mvarProvider = "SQLOLEDB"
    
    AryaSettingFile = "C:\Aryasetting.txt"

    Dim f As New FileSystemObject
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim IsFileExist As Boolean
    Dim LenghStr As Integer
        
   IsFileExist = filetemp.FileExists(AryaSettingFile)

   If IsFileExist = False Then
      MsgBox "Setting File Did Not Exist"
      End
   End If
   
    Set tempstring = filetemp.OpenTextFile(AryaSettingFile, ForReading, False, TristateFalse)
    
    Do While tempstring.AtEndOfLine = False
       Str = tempstring.ReadLine
       LenghStr = InStr(1, Str, "=", vbTextCompare)
       
       If InStr(1, Str, "ServerName", vbTextCompare) Then
          DataSource = CStr(Mid(Str, LenghStr + 1))
       
     ''''  ElseIf InStr(1, str, "RestoreDataBaseName", vbTextCompare) Then
       ElseIf InStr(1, Str, "DbName", vbTextCompare) Then
          DbName = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "RestoreDataBaseData", vbTextCompare) Then
          RestoreDataBaseData = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "RestoreDataBaseLog", vbTextCompare) Then
          RestoreDataBaseLog = CStr(Mid(Str, LenghStr + 1))
       
       End If
    Loop
    tempstring.Close
    If RestoreDataBaseData = "" Then
        RestoreDataBaseData = "D:\Arya\Data"
    End If
    If RestoreDataBaseLog = "" Then
        RestoreDataBaseLog = "D:\Arya\Data"
    End If
    
    If PosConnection.State = 1 Then PosConnection.Close
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    TxtDataSource.Text = DataSource
    TxtServer.Text = DataSource
    TxtDataPath.Text = RestoreDataBaseData & "\" & TxtDbName.Text & "_Data.mdf"
    TxtLogPath.Text = RestoreDataBaseLog & "\" & TxtDbName.Text & "_Log.ldf"
    Conn.Open "provider =  " + mvarProvider + "; data source = " + DataSource + ";Initial Catalog =msdb;user id=" + Trim(TxtUserName.Text) + ";Password=" + Trim(txtPassword.Text)
    
    formloadFlag = False
    
    FillDataBase
    
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

    SetKbLayout LANG_EN_US
 
End Sub

Private Sub Form_Resize()

'    On Error Resume Next
'    If Me.ScaleHeight > 0 Then
'       Me.Height = iHeight
'       Me.Width = iWidth
'    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    If strConnectionString = "" Then
        ShowMessage " œÌ «»Ì” «‰ Œ«» ‰‘œ, ”—Ê— Ê œÌ «»Ì” ÅÌ‘ ›—÷ «” ›«œÂ „Ì ‘Êœ.", True, False, " «∆Ìœ", ""
        strConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & clsArya.DBLogin & "; Password = " & SqlPass & " ; Initial Catalog = " & clsArya.DbName & ";Data Source=" & clsArya.ServerName
        AccstrConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & clsArya.DBLogin & "; Password = " & SqlPass & " ; Initial Catalog = " & clsArya.DbName & ";Data Source=" & clsArya.ServerName
        CrystallConnection = "Data Source=" & clsArya.ServerName & ";UID=" & clsArya.DBLogin & ";PWD=" & SqlPass & ";DSQ= " & clsArya.DbName & ";"
        
        Call ODBCSetting(clsArya.ServerName, clsArya.DbName)
        
    Else
        ShowMessage "‰«„ ”—Ê— : " & DataSource & Chr(10) & Chr(13) & "Ê ‰«„ »«‰ﬂ «ÿ·«⁄« Ì :" & DbName, True, False, " «∆Ìœ", ""
        strConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & clsArya.DBLogin & "; Password = " & SqlPass & " ; Initial Catalog = " & DbName & ";Data Source=" & DataSource
        AccstrConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & clsArya.DBLogin & "; Password = " & SqlPass & " ; Initial Catalog = " & DbName & ";Data Source=" & DataSource
        CrystallConnection = "Data Source=" & clsArya.ServerName & ";UID=" & clsArya.DBLogin & ";PWD=" & SqlPass & ";DSQ= " & DbName & ";"
        
        Call ODBCSetting(DataSource, DbName)
        
        ClsFormAccess.Class_Initialize
    
    End If
    
    If Conn.State = 1 Then Conn.Close
    Set Conn = Nothing
    
    
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
     
    CmdPreSet.Enabled = False
    
    SetKbLayout LANG_Pr_IR
    VarActForm = ""
End Sub

Private Sub FwbtnEsc_Click()
    Unload Me
End Sub

Private Sub FWBtnOK_Click()
Dim i As Integer
Dim lastbackid As Integer

Dim tem_str As Date

On Error GoTo ErrHandler

''''    Select Case LTrim(Right(txtFile, 3))
''''
''''
''''        Case "fbk", "Bak:"
        
''''            Conn.Open "provider =  " + mvarProvider + "; data source = " + DataSource + ";Initial Catalog =msdb;user id=sa;Password=" & SqlPass & ""
            
            ShowMessage "»«“ê—œ«‰Ì ›«Ì· " & Trim(txtFile.Text) & " —ÊÌ »«‰ﬂ «ÿ·«⁄« Ì " & Trim(TxtDbName.Text) & " . œ— ’Ê—  ÊÃÊœ »«‰ﬂ „“»Ê— «ÿ·«⁄«  »— —ÊÌ ¬‰ ‰Ê‘ Â „Ì ‘Êœ . ¬Ì« „ÿ„∆‰ Â” Ìœ ø", True, True, "»·Ì", "ŒÌ— "
            If mvarMsgIdx = vbNo Then Exit Sub
            If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
            If Conn.State = 1 Then Conn.Close
            Conn.Open "provider =  " + mvarProvider + "; data source = " + Trim(TxtServer.Text) + ";Initial Catalog =msdb;user id=" + Trim(TxtUserName.Text) + ";Password=" + Trim(txtPassword.Text)
            
            With cmd
                .ActiveConnection = Conn
                .CommandType = adCmdText
          '      .CommandText = "exec test"
                .CommandText = "RESTORE DATABASE " & Trim(TxtDbName.Text) & " FROM DISK = '" & Trim(txtFile.Text) & "' With FILE = 1,  NOUNLOAD ,  STATS = 10,  RECOVERY ,  REPLACE , "
                .CommandText = .CommandText & "Move N'Total_Data' TO N'" & Trim(TxtDataPath.Text) & "',"
                .CommandText = .CommandText & "Move N'Total_Log' TO N'" & Trim(TxtLogPath.Text) & "'"

            End With
            cmd.Execute
            cmd.Cancel

            Conn.Close
            MsgBox "»«“ê—œ«‰Ì »« „Ê›ﬁÌ  «Ã—« ‘œ "
       
''''       Case "dbk":
''''            Conn.Open "provider =  " + mvarProvider + "; data source = " + DataSource + ";Initial Catalog =msdb;user id=sa;Password=" & SqlPass & ""
'''''            tem_str = FileDateTime(txtFile.Text)
'''''            rs.Open " select IsNull(max(position),1) as lastid from backupset where database_name='pubs' and  DATEPART(year,backup_finish_date)=" & Year(tem_str) & " and  DATEPART(month,backup_finish_date)=" & Month(tem_str) & "  and DATEPART(day,backup_finish_date)=" & Day(tem_str), Conn
'''''            lastbackid = rs!lastid
'''''            If lastbackid > 1 Then
'''''               MsgBox " »—«Ì «Ì‰  «—ÌŒ " & lastbackid & "‰”ŒÂ Å‘ Ì»«‰ ÊÃÊœ œ«—œ Ê ¬Œ—Ì‰ ¬‰ »«“ê—œ«‰Ì „Ì ‘Êœ "
'''''            End If
'''''            rs.Close
''''            With cmd
''''                .ActiveConnection = Conn
''''                .CommandType = adCmdText
''''                .CommandText = "RESTORE DATABASE Pubs FROM DISK = '" & txtFile.Text & "' With Replace "
'''''                .CommandText = "RESTORE DATABASE Pubs FROM DISK = '" & txtFile.Text & "' With Replace,File=" & lastbackid
''''            End With
''''
''''            Conn.Open "provider =  " & mvarProvider & "; data source = " & DataSource & ";Initial Catalog = " & DbName & ";user id=sa;password=" & SqlPass & ";"
''''            Dim strFields As String
''''            Dim Rst As New ADODB.Recordset
''''            Rst.Open "select  top 1 * from tFacM", Conn, adOpenDynamic, adLockOptimistic, adCmdText
''''            strFields = ""
''''            For i = 0 To Rst.Fields.Count - 2
''''                strFields = strFields & ",[" & Rst.Fields(i).Name & "]"
''''            Next i
''''            If strFields <> "" Then
''''                strFields = Right(strFields, Len(strFields) - 1)
''''            End If
''''            Set Rst = Nothing
''''            Conn.Close
''''
''''            cmd.Execute
''''            cmd.Cancel
''''            Me.DeleteDb "tCust", "[Date]", DbName, "Pubs"
''''            Me.DeleteDb "tSupplier", "[Date]", DbName, "Pubs"
''''            Me.DeleteDb "tRepFacEditm", "[Date]", DbName, "Pubs"
''''            Me.DeleteDb "tFacD2", "[intserialNo]", DbName, "Pubs"
''''            Me.DeleteDb "tFacD", "[intserialNo]", DbName, "Pubs"
''''            Me.DeleteDb "tFacM", "[Date]", DbName, "Pubs"
''''            Me.InsertDb "tFacM", "[Date]", DbName, "Pubs", strFields
''''            Me.InsertDb "tFacD", "[intserialNo]", DbName, "Pubs"
''''            Me.InsertDb "tFacD2", "[intserialNo]", DbName, "Pubs"
''''            Me.InsertDb "tRepFacEditm", "[Date]", DbName, "Pubs"
''''            Me.InsertDb "tSupplier", "[Date]", DbName, "Pubs"
''''            Me.InsertDb "tCust", "[Date]", DbName, "Pubs"
''''            Me.DropDb "tCust", "Pubs"
''''            Me.DropDb "tSupplier", "Pubs"
''''            Me.DropDb "tRepFacEditm", "Pubs"
''''            Me.DropDb "tFacD2", "Pubs"
''''            Me.DropDb "tFacD", "Pubs"
''''            Me.DropDb "tFacM", "Pubs"
''''            Conn.Close
''''
''''          MsgBox "»«“ê—œ«‰Ì »« „Ê›ﬁÌ  «Ã—« ‘œ "
''''
''''       Case Else:
''''       MsgBox " ›«Ì· „Ê—œ ‰Ÿ— ÅÌœ« ‰‘œ"
''''    End Select
  '

Exit Sub

ErrHandler:
    ShowMessage err.Description, True, False, " «ÌÌœ", " "
    
    
End Sub

Public Sub ExitForm()
    Unload Me
End Sub

Public Sub InsertDb(mvarTable As String, mvarfldCon As String, dstDatabase As String, srcDatabase As String, Optional strFields As String)
On Error GoTo ErrorHandler
Dim Rst As New ADODB.Recordset

Rst.Open "select top 1 from " & mvarTable

With cmd
    .ActiveConnection = Conn
    .CommandType = adCmdText
    If strFields = "" Then
        .CommandText = " Insert Into " & dstDatabase & ".dbo." & mvarTable & " Select * From " & srcDatabase & ".dbo." & mvarTable & " Where " & mvarfldCon & " Not In (Select " & mvarfldCon & " From " & dstDatabase & ".dbo." & mvarTable & ")"
    Else
        .CommandText = " Insert Into " & dstDatabase & ".dbo." & mvarTable & " ( " & strFields & " ) " & "  Select  " & strFields & " From " & srcDatabase & ".dbo." & mvarTable & " Where " & mvarfldCon & " Not In (Select " & mvarfldCon & " From " & dstDatabase & ".dbo." & mvarTable & ")"
    End If
End With
cmd.Execute
cmd.Cancel
Exit Sub
ErrorHandler:
    Resume Next
End Sub

Public Sub DeleteDb(mvarTable As String, mvarfldCon As String, dstDatabase As String, srcDatabase As String)
On Error GoTo ErrorHandler
With cmd
    .ActiveConnection = Conn
    .CommandType = adCmdText
    .CommandText = "Delete  From " & dstDatabase & ".dbo." & mvarTable & " Where " & mvarfldCon & " In (Select " & mvarfldCon & " From " & srcDatabase & ".dbo." & mvarTable & ")"
End With
cmd.Execute
cmd.Cancel
Exit Sub
ErrorHandler:
    Resume Next
End Sub

Public Sub DropDb(mvarTable As String, Database As String)
On Error GoTo ErrorHandler
With cmd
    .ActiveConnection = Conn
    .CommandType = adCmdText
    .CommandText = " Drop Table " & Database & ".dbo." & mvarTable
End With
cmd.Execute
cmd.Cancel
Exit Sub
ErrorHandler:
    Resume Next
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub TxtDataSource_Change()
    cmbDataBase.Clear
End Sub

Private Sub TxtDbName_Change()
    TxtDataPath.Text = RestoreDataBaseData & "\" & TxtDbName.Text & "_Data.mdf"
    TxtLogPath.Text = RestoreDataBaseLog & "\" & TxtDbName.Text & "_Log.ldf"
End Sub


