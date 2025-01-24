VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private ClsCnn As New ADODB.Connection
Private cmd As New ADODB.Command
Private Rst As New ADODB.Recordset
Private RstTmp As New ADODB.Recordset
Private mvarProvider As String
Private mvarDatasource As String
Private mvarInitialCatalog As String
Private mvarTable As String
Private mvarOrder As String
Private mvarText As String
Private mvarUserID As String
Private mvarReadOnly As Boolean
Private mvarAttachType As Variant
Private mvarJoin As Boolean
Private mvarJoinClient As Boolean
Private mvarJoinServer As Boolean
Private mvarTimeOut As Integer
Public Enum EnumAlign
    vbLeft = 1
    vbRight = 2
    vbCenter = 3
End Enum

Public Function Comparing(String1 As String, String2 As String) As Boolean
If Format(Trim(String1), ">") = Format(Trim(String2), ">") Then
    Comparing = True
Else
    Comparing = False
End If
End Function

Public Property Let AttachType(ByVal vData As Variant)
    mvarAttachType = vData
End Property

Public Property Get AttachType() As Variant
    AttachType = mvarAttachType
End Property

Public Property Let UserID(ByVal vData As String)
    mvarUserID = vData
End Property

Public Property Get UserID() As String
    UserID = mvarUserID
End Property


Public Property Let InitialCatalog(ByVal vData As String)
    mvarInitialCatalog = vData
End Property

Public Property Get InitialCatalog() As String
    InitialCatalog = mvarInitialCatalog
End Property

Public Property Let DataSource(ByVal vData As String)
    mvarDatasource = vData
End Property

Public Property Get DataSource() As String
    DataSource = mvarDatasource
End Property

Public Property Let Provider(ByVal vData As String)
    mvarProvider = vData
End Property

Public Property Get Provider() As String
    Provider = mvarProvider
End Property

Public Property Let Table(ByVal vData As String) '‰«„ table „œ‰Ÿ— —« œ—Ì«›  „Ìœ«—œ
    mvarTable = vData
End Property

Public Property Get Table() As String '‰«„ table —« ⁄Êœ  „ÌœÂœ
    Table = mvarTable
End Property

Public Property Let Order(ByVal vData As String) '‰«„ table „œ‰Ÿ— —« œ—Ì«›  „Ìœ«—œ
    mvarOrder = vData
End Property

Public Property Get Order() As String '‰«„ table —« ⁄Êœ  „ÌœÂœ
    Order = mvarOrder
End Property

Public Property Let ReadOnly(ByVal vData As Boolean) '»«‰ﬂ —« œ— „œ readonly ›⁄«· „Ì ﬂ‰œ
    mvarReadOnly = vData
End Property

Public Property Get ReadOnly() As Boolean 'Ê÷⁄Ì  »«‰ﬂ —« «“ ‰Ÿ— readonly »Êœ‰ «—Ã«⁄ „Ì œÂœ
    ReadOnly = mvarReadOnly
End Property

Public Property Let JoinClient(mydata As Boolean) '«Ì‰ Œ«’Ì  „ﬁ«œÌ— table —« œ— datagrid «⁄„«· „Ì ﬂ‰œ
If mydata Then
    Me.Connect = True
    With Rst
        .ActiveConnection = ClsCnn
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        If Not mvarReadOnly Then
            .LockType = adLockOptimistic
        Else
            .LockType = adLockReadOnly
        End If
        .Source = mvarTable
    End With
    mvarJoinClient = True
Else
    mvarJoinClient = False
End If
Exit Property
ErrorHandler:
    MsgBox "«— »«ÿ »« »«‰ﬂ »—ﬁ—«— ‰‘œ"
End Property

Public Property Get JoinClient() As Boolean '«Ì‰ Œ«’Ì  „‘Œ’ „Ì ﬂ‰œ ¬Ì« «— »«ÿ »« »«‰ﬂ »—ﬁ—«— ‘œÂ «”  Ì« ŒÌ—
    JoinClient = mvarJoinClient
End Property

Public Property Let join(mydata As Boolean) '«Ì‰ Œ«’Ì  „ﬁ«œÌ— table —« œ— datagrid «⁄„«· „Ì ﬂ‰œ
If mydata Then
    Me.Connect = True
   On Error GoTo ErrorHandler
    
    If Rst.State <> 0 Then
        Rst.Close
    End If
    With Rst
        .ActiveConnection = ClsCnn
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        If Not mvarReadOnly Then
            .LockType = adLockOptimistic
        Else
            .LockType = adLockReadOnly
        End If
        If IsEmpty(mvarOrder) Or Trim(mvarOrder) = "" Then
            .Source = mvarTable
        Else
            .Source = "Select * From (" & mvarTable & ") Order By " & mvarOrder
        End If
    End With
    mvarJoin = True
Else
    mvarJoin = False
End If
Exit Property
ErrorHandler:
    MsgBox "«— »«ÿ »« »«‰ﬂ »—ﬁ—«— ‰‘œ"
End Property

Public Property Get join() As Boolean '«Ì‰ Œ«’Ì  „‘Œ’ „Ì ﬂ‰œ ¬Ì« «— »«ÿ »« »«‰ﬂ »—ﬁ—«— ‘œÂ «”  Ì« ŒÌ—
    join = mvarJoin
End Property

Public Property Let JoinServer(mydata As Boolean) '«Ì‰ Œ«’Ì  „ﬁ«œÌ— table —« œ— datagrid «⁄„«· „Ì ﬂ‰œ
If mydata Then
    Me.Connect = True
   On Error GoTo ErrorHandler
    If Rst.State <> 0 Then
        Rst.Close
    End If
    With Rst
        .ActiveConnection = ClsCnn
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        If Not mvarReadOnly Then
            .LockType = adLockPessimistic
        Else
            .LockType = adLockReadOnly
        End If
        .Source = mvarTable
        .Open
    End With
    mvarJoinServer = True
Else
    mvarJoinServer = False
End If
Exit Property
ErrorHandler:
    MsgBox "«— »«ÿ »« »«‰ﬂ »—ﬁ—«— ‰‘œ"
End Property

Public Property Get JoinServer() As Boolean '«Ì‰ Œ«’Ì  „‘Œ’ „Ì ﬂ‰œ ¬Ì« «— »«ÿ »« »«‰ﬂ »—ﬁ—«— ‘œÂ «”  Ì« ŒÌ—
    join = mvarJoinServer
End Property

Public Property Let SetMenu(User As Variant) 'ﬂ«—»— «” ›«œÂ ﬂ‰‰œÂ —« „⁄Ì‰ „Ì ‰„«Ìœ Ê »— «”«” ¬‰ „‰ÊÂ« ›⁄«· „Ìê—œ‰œ
Dim obj As Object
On Error Resume Next
For Each obj In mdifrm
    If TypeOf obj Is Menu Then
        If obj.Caption <> "-" Then
            With cmd
                .ActiveConnection = ClsCnn
                .CommandType = adCmdText
                .CommandText = "Select " & obj.Name & "  From Access where Username = " & "'" & User & "'"
            End With
            Set RstTmp = cmd.Execute
            If Trim(RstTmp.Fields(obj.Name).Value) = "⁄œ„ œ” —”Ì" Or Left(Trim(RstTmp.Fields(0).Value), 3) = "⁄œ„" Then
                obj.Enabled = False
            Else
                obj.Enabled = True
            End If
            RstTmp.Close
        End If
    End If
Next
'Cmd.Cancle
Exit Property
ErrorHandler:
    MsgBox "“Ì—„‰Ê ﬁ«»·Ì  ⁄œ„ —ƒÌ  ‰œ«—œ"
End Property

Public Property Get Recordset() As ADODB.Recordset
On Error GoTo ErrorHandler 'Resume Next '
If Rst.State = adStateClosed Then
    Rst.Open
End If
    
Set Recordset = Rst
Exit Property
ErrorHandler:
    Exit Property
End Property

Public Property Let Connect(mydata As Boolean)
If mydata Then
    If ClsCnn.State = adStateOpen Then
        ClsCnn.Close
    End If
         ClsCnn.Open strConnectionString
End If
Me.TimeOut = 2
Exit Property
ErrorHandler:
    Exit Property
End Property

Public Property Get Connection() As ADODB.Connection
    Set Connection = ClsCnn
End Property

Public Property Let TimeOut(vData As Integer)
    If ClsCnn.State = adStateOpen Then
       ClsCnn.CommandTimeout = vData
    End If
    mvarTimeOut = vData
End Property

Public Property Get TimeOut() As Integer
    TimeOut = mvarTimeOut
End Property

Public Property Let Text(mydata As String)
    mvarText = mydata
End Property

Public Property Get Text() As String
    Text = mvarText
End Property

Public Sub FillList(obj As Object, Rc As ADODB.Recordset, Code As String, Name As String, Optional Name2 As String)
Dim ii As Integer
On Error Resume Next
       obj.Clear
       Rc.moveFirst
       ii = 0
       Do While Not Rc.EOF()
          If Not IsNull(Rc.Fields(Name)) Then
              obj.AddItem Rc.Fields(Name)
'              If Left(Rc.Fields(Code), 1) <> "0" Then
                  obj.ItemData(ii) = Rc.Fields(Code)
'              Else
'                  obj.ItemData(ii) = "9999" & Rc.Fields(Code)
'              End If
              ii = ii + 1
          End If
          Rc.MoveNext
       Loop
'       obj.ListIndex = 0
End Sub

Private Sub Class_Terminate()
    Set ClsCnn = Nothing
    Set cmd = Nothing
    Set Rst = Nothing
    Set RstTmp = Nothing
End Sub
