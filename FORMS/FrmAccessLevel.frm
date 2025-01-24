VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAccessLevel 
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   Icon            =   "FrmAccessLevel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   9135
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00008000&
      Caption         =   " «ÌÌœ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   6825
      Width           =   2445
   End
   Begin VB.CommandButton cmdSelectAll 
      BackColor       =   &H000080FF&
      Caption         =   "«‰ Œ«» Â„Â"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Tag             =   "1"
      Top             =   5850
      Width           =   2445
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00008000&
      Caption         =   "«÷«›Â"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1095
      Width           =   765
   End
   Begin VB.ListBox lstAccessLevel 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   5400
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1740
      Width           =   3495
   End
   Begin VB.TextBox txtAccessLevel 
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
      Height          =   525
      Left            =   6150
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   2745
   End
   Begin MSComctlLib.TreeView trObjects 
      Height          =   8655
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   15266
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "FrmAccessLevel.frx":A4C2
      TabIndex        =   8
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”ÿÊÕ œ” —”Ì"
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
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "”ÿÊÕ œ” —”Ì"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   6150
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   630
      Width           =   2745
   End
End
Attribute VB_Name = "FrmAccessLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDate As New clsDate
Dim i As Integer
Dim strMaxAccess As String
Dim strNewAccess As String
Dim Parameter() As Parameter
Dim CurrentLevel As Integer

Private Sub cmdAdd_Click()
  
    If Trim(txtAccessLevel.Text) <> "" Then
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@Description", adVarWChar, 50, Trim(txtAccessLevel.Text))
        RunParametricStoredProcedure "InsertAccessLevel", Parameter
    
    End If
    FillLstAccessLevel
    
End Sub


Private Sub cmdUpdate_Click()

    Dim ObjectCode As String
    
    If lstAccessLevel.SelCount < 1 Then Exit Sub
    
    strNewAccess = ""
    For i = 1 To trObjects.Nodes.Count
        If trObjects.Nodes.Item(i).Checked = True Then
            strNewAccess = strNewAccess & "1"
        Else
            strNewAccess = strNewAccess & "0"
        End If
    Next i
    
'    Dim strFinalAccess As String
'    strFinalAccess = ""
'
'    For i = 1 To Len(strMaxAccess)
'        strFinalAccess = strFinalAccess & CStr(Val(Mid(strMaxAccess, i, 1)) And Val(Mid(strNewAccess, i, 1)))
'    Next i
    
    For i = 1 To trObjects.Nodes.Count
        If trObjects.Nodes(i).Checked = True And CStr(Val(Mid(strMaxAccess, i, 1)) And Val(Mid(strNewAccess, i, 1))) = "1" Then
            If InStr(1, trObjects.Nodes(i).FullPath, "/") > 0 Then
                If trObjects.Nodes(i).Parent.Checked = True Then
                    ObjectCode = ObjectCode & Left(trObjects.Nodes(i).Key, Len(trObjects.Nodes(i).Key) - 1)
                End If
            Else
                ObjectCode = ObjectCode & Left(trObjects.Nodes(i).Key, Len(trObjects.Nodes(i).Key) - 1) & ","
            End If
        End If
    Next i
    
    If Len(ObjectCode) >= 1 Then
       ObjectCode = Left(ObjectCode, Len(ObjectCode) - 1)
    End If
    ReDim Parameter(1) As Parameter
    For i = 0 To lstAccessLevel.ListCount - 1
        If lstAccessLevel.Selected(i) = True Then
            If lstAccessLevel.ItemData(i) > mVarAccessLevel Then
                Parameter(0) = GenerateInputParameter("@intAccessLevel", adInteger, 4, lstAccessLevel.ItemData(i))
                Exit For
            Else
                frmMsg.fwlblMsg.Caption = "‘„« ‰„Ì  Ê«‰Ìœ «Ì‰ ”ÿÕ œ” —”Ì —«  €ÌÌ— œÂÌœ"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                Exit Sub
            End If
        End If
    Next i
    
    Parameter(1) = GenerateInputParameter("@ObjectCode", adVarWChar, 4000, ObjectCode)
    RunParametricStoredProcedure "Update_tAccess_Object", Parameter
    
    RefreshtrObjects
    
End Sub

Private Sub cmdSelectAll_Click()
    If trObjects.Nodes.Count > 0 Then
        If cmdSelectAll.Tag = 1 Then
            cmdSelectAll.Tag = 0
            For i = 1 To trObjects.Nodes.Count
                trObjects.Nodes(i).Checked = True
            Next i
            cmdSelectAll.Caption = "Å«ò ò—œ‰ Â„Â"
        Else
            cmdSelectAll.Tag = 1
            For i = 1 To trObjects.Nodes.Count
                trObjects.Nodes(i).Checked = False
            Next i
            cmdSelectAll.Caption = "«‰ Œ«» Â„Â"
        End If
    End If

End Sub

Private Sub Form_Activate()
    Dim i As Integer
    
    AllButton vbOff, True
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    VarActForm = Me.Name
    
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

    If ClsFormAccess.frmFacRecursive = False Then
        Unload Me
        Exit Sub
    End If

    CenterTop Me
    
    VarActForm = Me.Name
            
    FillLstAccessLevel
    FilltrObjects
    For i = 0 To lstAccessLevel.ListCount - 1
        If lstAccessLevel.ItemData(i) = mVarAccessLevel Then
            lstAccessLevel.Selected(i) = True
        End If
    Next i
    RefreshtrObjects
    GenerateAccessString strMaxAccess
    
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


    
End Sub
Public Sub ExitForm()

    Unload Me

End Sub

Private Sub FillLstAccessLevel()

    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set Rst = RunParametricStoredProcedure2Rec("GetAccessLevel", Parameter)
    lstAccessLevel.Clear
    If Rst.EOF <> True And Rst.BOF <> True Then
        While Rst.EOF <> True
            lstAccessLevel.AddItem Rst.Fields("Description").Value
            lstAccessLevel.ItemData(lstAccessLevel.ListCount - 1) = Rst.Fields("intAccessLevel").Value
            Rst.MoveNext
        Wend
    
    End If
    
    Set Rst = Nothing
    
End Sub
Private Sub FilltrObjects()

    Dim Rst As New ADODB.Recordset
    ReDim Parameter(1) As Parameter
    Dim varNode As node

    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    If lstAccessLevel.SelCount = 0 Then
        Parameter(1) = GenerateInputParameter("@intAccessLevel", adInteger, 4, 0)
    Else
        For i = 0 To lstAccessLevel.ListCount - 1
            If lstAccessLevel.Selected(i) = True Then
                Parameter(1) = GenerateInputParameter("@intAccessLevel", adInteger, 4, lstAccessLevel.ItemData(i))
                Exit For
            End If
        Next i
    End If
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Object_Access", Parameter)

    trObjects.Nodes.Clear
    
    While Rst.EOF <> True
        If Rst.Fields("ObjectParent").Value = 0 Then
            Set varNode = trObjects.Nodes.Add(, , CStr(Rst.Fields("intObjectCode").Value) & "n", Rst.Fields("ObjectName").Value)
        Else
            Set varNode = trObjects.Nodes.Add(CStr(Rst.Fields("ObjectParent").Value) & "n", tvwChild, CStr(Rst.Fields("intObjectCode").Value) & "n", Rst.Fields("ObjectName").Value)
        End If
        If IsNull(Rst.Fields("intAccessLevel")) <> True Then
            varNode.Checked = True
        Else
            varNode.Checked = False
        End If
        Rst.MoveNext
    Wend
    Set Rst = Nothing
End Sub

Private Sub RefreshtrObjects()

    Dim Rst As New ADODB.Recordset
    ReDim Parameter(1) As Parameter
    Dim varNode As node

    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    If lstAccessLevel.SelCount = 0 Then
        Parameter(1) = GenerateInputParameter("@intAccessLevel", adInteger, 4, 0)
    Else
        For i = 0 To lstAccessLevel.ListCount - 1
            If lstAccessLevel.Selected(i) = True Then
                Parameter(1) = GenerateInputParameter("@intAccessLevel", adInteger, 4, lstAccessLevel.ItemData(i))
                Exit For
            End If
        Next i
    End If
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Object_Access", Parameter)
    
    While Rst.EOF <> True
    
        Set varNode = trObjects.Nodes(CStr(Rst.Fields("intObjectCode").Value) & "n")
        If IsNull(Rst.Fields("intAccessLevel")) <> True Then
            varNode.Checked = True
        Else
            varNode.Checked = False
        End If
        Rst.MoveNext
        
    Wend
    Set Rst = Nothing
    
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Private Sub Form_Unload(Cancel As Integer)

    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    VarActForm = ""
    
    Dim Obj As Object
    For Each Obj In Forms
        If TypeOf Obj Is Form Then
            If Obj.Name <> "mdifrm" And Obj.Name <> Me.Name And Obj.Name <> "frmAbout" Then
                Obj.Show
            End If
        End If

    Next Obj
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub

Private Sub lstAccessLevel_ItemCheck(Item As Integer)
    If lstAccessLevel.ItemData(Item) >= mVarAccessLevel Then
        If lstAccessLevel.SelCount > 1 Then
            For i = 0 To lstAccessLevel.ListCount - 1
                If i <> Item Then
                    lstAccessLevel.Selected(i) = False
                End If
            Next i
        End If
        RefreshtrObjects
    Else
        lstAccessLevel.Selected(Item) = False
        frmMsg.fwlblMsg.Caption = "‘„« ‰„Ì  Ê«‰Ìœ «Ì‰ ”ÿÕ œ” —”Ì —«  €ÌÌ— œÂÌœ"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
    End If
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub trObjects_NodeCheck(ByVal node As MSComctlLib.node)
    Dim varNode As node
    If node.Children < 1 Then Exit Sub
    Set varNode = node.Child
    varNode.Checked = node.Checked
    Set varNode = varNode.Next
    
    If node.Children > 1 Then
        For i = 2 To node.Children
            varNode.Checked = node.Checked
            Set varNode = varNode.Next
        Next i
    End If
End Sub

Private Sub GenerateAccessString(ByRef Str As String)
    
    Str = ""
    For i = 1 To trObjects.Nodes.Count
        If trObjects.Nodes.Item(i).Checked = True Then
            Str = Str & "1"
        Else
            Str = Str & "0"
        End If
    Next i

End Sub
