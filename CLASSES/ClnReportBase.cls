VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClnReportBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClsReportBase"
Attribute VB_Ext_KEY = "Member0" ,"ClsReportBase"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mvarRc As New ADODB.Recordset
Private ClsGl As New ClsGl
Private ObjNewMember As New ClsReportBase

Public Function Add(DscrFull As String, DscrIn As String, DscrOut As String, Optional Frmt As String, Optional Key As String, Optional DataType As String) As ClsReportBase
    ObjNewMember.DscrFull = DscrFull
    ObjNewMember.DscrIn = DscrIn
    ObjNewMember.DscrOut = DscrOut
    ObjNewMember.Frmt = Frmt
    ObjNewMember.Key = Key
    ObjNewMember.DataType = DataType
    If Len(Key) = 0 Then
        mcol.Add ObjNewMember
    Else
        mcol.Add ObjNewMember, Key
    End If
    Set Add = ObjNewMember
    Set ObjNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClsReportBase
Attribute Item.VB_UserMemId = 0
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property


Public Sub remove(vntIndexKey As Variant)

    mcol.remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcol.[_NewEnum]
End Property


Private Sub Class_Initialize()
'creates the collection when this class is created
Set mcol = New Collection
Me.Add "�����", "�� �����", "�� �����", , "Dat", "String"
Me.Add "������", "�� ������", "�� ������", , "Fac", "Integer"
Me.Add "�����", "�� �����", "�� �����", , "StationId", "Integer"
Me.Add "�����", "�� ������", "�� ������", , "Customer", "Integer"
Me.Add "����� �����", "�� ����� �����", "�� ����� �����", , "Supplier", "Integer"
Me.Add "�����", "�� �����", "�� �����", , "User", "String"
Me.Add "����", "�� ����", "�� ����", , "Time", "String"
Me.Add "�����", "�������", "�� �����", , "SDiscount", "Double"
Me.Add "����", "������", "������", , "SPrice", "Double"
Me.Add "���", "�ҁ��", "�ǁ��", , "Payk", "Integer"
Me.Add "�����", "�������", "�������", , "SerialNo", "Double"
Me.Add "���� ����", "�Ґ��� ����", "�ǐ��� ����", , "L1", "String"
Me.Add "���� ����", "�Ґ��� ����", "�ǐ��� ����", , "L2", "String"
Me.Add "��� ����", "����� ����", "����� ����", , "Good", "String"
Me.Add "����", "������", "������", , "Shift", "String"
Me.Add "������", "�� ������", "�� ������", , "garson", "Integer"
Me.Add "��", "�� ��", "�� ��", , "carrier", "Integer"
Me.Add "����� �ǘ���", "����� �ǘ���", "����� �ǘ���", , "FactorNo", "Integer"
Me.Add "���� �ǁ", "���� �ǁ", "���� �ǁ", , "PrintFormat", "Integer"
Me.Add "�����", "�� �����", "�� �����", , "Description", "String"
Me.Add "�����", "�� ", "�� ", , "StoreDescription", "String"
Me.Add "����", "��", "�� ", , "CalculatedGood", "String"
Me.Add "��� ����", "��", "�� ", , "CustomerType", "String"
Me.Add "������", "��", "�� ", , "Details", "String"
Me.Add "���� ����", "��", "�� ", , "PaperType", "String"
Me.Add "��� ����", "��� ����", " ", , "AccountYear", "Integer"
Me.Add "����", "����", " ", , "Branch", "String"
Me.Add "���", "���", " ", , "Status", "String"
Me.Add "���", "�� ���", "�� ��� ", , "Table", "Integer"
Me.Add "�����", "��", "�� ", , "SortOrder", "String"
Me.Add "���", "�� ���", "�� ��� ", , "Month", "Integer"
Me.Add "�������", "�� �������", "�� �������", , "Seller", "Integer"
Me.Add "����Ԑ�", "�� ����Ԑ�", "�� ����Ԑ�", , "HairDresser", "Integer"
Me.Add "��� �������", "��", "�� ", , "SellerReportType", "String"
Me.Add "����� �����", "��", "�� ", , "OrderKind", "String"
Me.Add "���� �����", "��", "�� ", , "ShowKind", "String"
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    On Error Resume Next
    Set mvarRc = Nothing
    Set ClsGl = Nothing
    Set ObjNewMember = Nothing
    Set mcol = Nothing
'    mdifrm.fwBtnCtrl.SetFocus
End Sub



