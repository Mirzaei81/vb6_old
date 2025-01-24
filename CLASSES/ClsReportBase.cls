VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsReportBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Collection1"
Option Explicit
Private mvarNum As Integer
Private mvarDscrFull As String 'local copy
Private mvarDscrIn As String 'local copy
Private mvarDscrOut As String 'local copy
Private mvarMinValue As Variant 'local copy
Private mvarMaxValue As Variant 'local copy
Private mvarFrmt As String 'local copy
Private mvarKey As String 'local copy
Private mvarPriority As Integer
Private mvarFieldName As String
Private mvarDataType As String
Private mvarDefaultMin As Variant
Private mvarDefaultMax As Variant
'local variable(s) to hold property value(s)
Private mvarFieldCode As String 'local copy
Private mvarFieldDscr As String 'local copy
Private mvarrcTmp As String 'local copy

Public Property Let rctmp(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.rcTmp = 5
    mvarrcTmp = vData
End Property

Public Property Get rctmp() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.rcTmp
    rctmp = mvarrcTmp
End Property



Public Property Let FieldDscr(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldDscr = 5
    mvarFieldDscr = vData
End Property


Public Property Get FieldDscr() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldDscr
    FieldDscr = mvarFieldDscr
End Property



Public Property Let FieldCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldCode = 5
    mvarFieldCode = vData
End Property


Public Property Get FieldCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldCode
    FieldCode = mvarFieldCode
End Property




Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Let Frmt(ByVal vData As String)
    mvarFrmt = vData
End Property

Public Property Get Frmt() As String
    Frmt = mvarFrmt
End Property

Public Property Let MaxValue(ByVal vData As Variant)
    mvarMaxValue = vData
End Property

Public Property Set MaxValue(ByVal vData As Variant)
    Set mvarMaxValue = vData
End Property

Public Property Get MaxValue() As Variant
    If IsObject(mvarMaxValue) Then
        Set MaxValue = mvarMaxValue
    Else
        MaxValue = mvarMaxValue
    End If
End Property

Public Property Let MinValue(ByVal vData As Variant)
    mvarMinValue = vData
End Property

Public Property Set MinValue(ByVal vData As Variant)
    Set mvarMinValue = vData
End Property

Public Property Get MinValue() As Variant
    If IsObject(mvarMinValue) Then
        Set MinValue = mvarMinValue
    Else
        MinValue = mvarMinValue
    End If
End Property

Public Property Let DscrOut(ByVal vData As String)
    mvarDscrOut = vData
End Property

Public Property Get DscrOut() As String
    DscrOut = mvarDscrOut
End Property

Public Property Let DscrIn(ByVal vData As String)
    mvarDscrIn = vData
End Property

Public Property Get DscrIn() As String
    DscrIn = mvarDscrIn
End Property

Public Property Let DscrFull(ByVal vData As String)
    mvarDscrFull = vData
End Property

Public Property Get DscrFull() As String
    DscrFull = mvarDscrFull
End Property

Public Property Let Num(mydata As Integer)
    mvarNum = mydata
End Property

Public Property Get Num() As Integer
    Num = mvarNum
End Property

Public Property Let Priority(mydata As Integer)
    mvarPriority = mydata
End Property

Public Property Get Priority() As Integer
    Priority = mvarPriority
End Property

Public Property Let FieldName(mydata As String)
    mvarFieldName = mydata
End Property

Public Property Get FieldName() As String
    FieldName = mvarFieldName
End Property

Public Property Let DefaultMin(mydata As Variant)
    mvarDefaultMin = mydata
End Property

Public Property Get DefaultMin() As Variant
    DefaultMin = mvarDefaultMin
End Property

Public Property Let DefaultMax(mydata As Variant)
    mvarDefaultMax = mydata
End Property

Public Property Get DefaultMax() As Variant
    DefaultMax = mvarDefaultMax
End Property

Public Property Let DataType(mydata As String)
    mvarDataType = mydata
End Property

Public Property Get DataType() As String
    DataType = mvarDataType
End Property

