VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicAccounting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_PerssonelAtf As Long
Private m_CustomerAtf As Long
Private m_SupplierAtf As Long
Private m_CosumerCompanyAtf As Long
Private m_SupplierCompanyAtf As Long

Public Property Get PerssonelAtf() As Long
    PerssonelAtf = m_PerssonelAtf
End Property
Public Property Let PerssonelAtf(ByVal vData As Long)
    m_PerssonelAtf = vData
End Property
Public Property Get CustomerAtf() As Long
    CustomerAtf = m_CustomerAtf
End Property
Public Property Let CustomerAtf(ByVal vData As Long)
    m_CustomerAtf = vData
End Property
Public Property Let SupplierAtf(ByVal vData As Long)
    m_SupplierAtf = vData
End Property
Public Property Get SupplierAtf() As Long
    SupplierAtf = m_SupplierAtf
End Property
Public Property Let CosumerCompanyAtf(ByVal vData As Long)
    m_CosumerCompanyAtf = vData
End Property
Public Property Get CosumerCompanyAtf() As Long
    CosumerCompanyAtf = m_CosumerCompanyAtf
End Property
Public Property Get SupplierCompanyAtf() As Long
    SupplierCompanyAtf = m_CosumerCompanyAtf
End Property
Public Property Let SupplierCompanyAtf(ByVal vData As Long)
    m_SupplierCompanyAtf = vData
End Property
Private Sub Class_Initialize()
    Dim file As New FileSystemObject
    Dim strReader As TextStream
    Dim strRead As String
    Dim indexEqual As Long
    
    If AccountingSettingFile = "" Then End    'Only  For  Make Exe File
    Set strReader = file.OpenTextFile(AccountingSettingFile, ForReading, False, TristateFalse)
    
    Do While strReader.AtEndOfLine = False
        strRead = strReader.ReadLine()
        indexEqual = InStr(1, strRead, "=", vbTextCompare)
        
        If InStr(1, strRead, "PerssonelAtf", vbTextCompare) Then
            m_PerssonelAtf = Mid(strRead, indexEqual + 1)
            
        ElseIf InStr(1, strRead, "CustomerAtf", vbTextCompare) Then
            m_CustomerAtf = Mid(strRead, indexEqual + 1)
        
        ElseIf InStr(1, strRead, "SupplierAtf", vbTextCompare) Then
            m_SupplierAtf = Mid(strRead, indexEqual + 1)
            
        ElseIf InStr(1, strRead, "ConsumerCompanyAtf", vbTextCompare) Then
            m_CosumerCompanyAtf = Mid(strRead, indexEqual + 1)
            
        ElseIf InStr(1, strRead, "SupplierCompanyAtf", vbTextCompare) Then
            m_SupplierCompanyAtf = Mid(strRead, indexEqual + 1)
            
        End If
    
    
    Loop
End Sub

