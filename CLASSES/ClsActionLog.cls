VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsActionLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Parameter() As Parameter

Private m_InsertFactor As Boolean
Private m_EditFactor As Boolean
Private m_RefferFactor As Boolean
Private m_GiveFactorToPayk As Boolean
Private m_PayCustomerFactor As Boolean
Private m_payPaykFactor As Boolean
Private m_PayGarsonFactor As Boolean
Private m_PaySellerFactor As Boolean
Private m_UnReferFactor As Boolean
Private m_RefferFromPaykAccount As Boolean
Private m_RefferFromSellerAccount As Boolean
Private m_RefferFromGarsonAccount As Boolean
Private m_DataSentToPos As Boolean
Private m_PosAcknowledged As Boolean
Private m_MovePaykToCustomCredit As Boolean
Private m_GiveFoodtoCustomer As Boolean
Private m_GiveFoodtoDeliverer As Boolean
Private m_MoveTableToCustomCredit As Boolean

Public Enum EnumActionLog

    InsertFactor = 1
    EditFactor = 2
    RefferFactor = 3
    GiveFactorToPayk = 4
    PayCustomerFactor = 5
    PayPaykFactor = 6
    PaySellerFactor = 7
    PayGarsonFactor = 7
    UnReferFactor = 8
    RefferFromPaykAccount = 9
    RefferFromSellerAccount = 10
    RefferFromGarsonAccount = 10
    DataSentToPos = 11
    PosAcknowledged = 12
    MovePaykToCustomCredit = 13
    GiveFoodtoCustomer = 14
    GiveFoodtoDeliverer = 15
    Printing = 16
    Reprint = 17
    InvoicePrint = 18
    MoveTableToCustomCredit = 19
    
End Enum

Public Property Get LogInsertFactor()
    LogInsertFactor = m_InsertFactor
End Property

Public Property Get LogEditFactor()
    LogEditFactor = m_EditFactor
End Property

Public Property Get LogRefferFactor()
    LogRefferFactor = m_RefferFactor
End Property

Public Property Get LogGiveFactorToPayk()
    LogGiveFactorToPayk = m_GiveFactorToPayk
End Property

Public Property Get LogPayCustomerFactor()
    LogPayCustomerFactor = m_PayCustomerFactor
End Property

Public Property Get LogPayPaykFactor()
    LogPayPaykFactor = m_payPaykFactor
End Property

Public Property Get LogPaySellerFactor()
    LogPaySellerFactor = m_PaySellerFactor
End Property
Public Property Get LogPayGarsonFactor()
    LogPayGarsonFactor = m_PayGarsonFactor
End Property

Public Property Get LogUnReferFactor()
    LogUnReferFactor = m_UnReferFactor
End Property

Public Property Get LogRefferFromPaykAccount()
    LogRefferFromPaykAccount = m_RefferFromPaykAccount
End Property

Public Property Get LogRefferFromSellerAccount()
    LogRefferFromSellerAccount = m_RefferFromSellerAccount
End Property
Public Property Get LogRefferFromGarsonAccount()
    LogRefferFromGarsonAccount = m_RefferFromGarsonAccount
End Property

Public Property Get LogDataSentToPos()
    LogDataSentToPos = m_DataSentToPos
End Property

Public Property Get LogPosAcknowledged()
    LogPosAcknowledged = m_PosAcknowledged
End Property

Public Property Get LogMovePaykToCustomCredit()
    LogMovePaykToCustomCredit = m_MovePaykToCustomCredit
End Property
Public Property Get LogMoveTableToCustomCredit()
    LogMoveTableToCustomCredit = m_MoveTableToCustomCredit
End Property

Private Sub Class_Initialize()

    Dim ClsCnn As New ADODB.Connection
    Dim Rst As New ADODB.Recordset
    
    ClsCnn.ConnectionString = strConnectionString
    ClsCnn.Open
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set Rst = RunParametricStoredProcedure2Rec("Get_Action", Parameter, ClsCnn)

    While Rst.EOF <> True
    
        Select Case Rst.Fields("ActionCode").Value
            Case EnumActionLog.EditFactor
                m_EditFactor = True
            Case EnumActionLog.GiveFactorToPayk
                m_GiveFactorToPayk = True
            Case EnumActionLog.InsertFactor
                m_InsertFactor = True
            Case EnumActionLog.PayCustomerFactor
                m_PayCustomerFactor = True
            Case EnumActionLog.PayGarsonFactor
                m_PayGarsonFactor = True
            Case EnumActionLog.PaySellerFactor
                m_PaySellerFactor = True
            Case EnumActionLog.PayPaykFactor
                m_payPaykFactor = True
            Case EnumActionLog.RefferFactor
                m_RefferFactor = True
            Case EnumActionLog.UnReferFactor
                m_UnReferFactor = True
            Case EnumActionLog.RefferFromPaykAccount
                m_RefferFromPaykAccount = True
            Case EnumActionLog.RefferFromGarsonAccount
                m_RefferFromGarsonAccount = True
            Case EnumActionLog.RefferFromSellerAccount
                m_RefferFromSellerAccount = True
            Case EnumActionLog.DataSentToPos
                m_DataSentToPos = True
            Case EnumActionLog.MovePaykToCustomCredit
                m_MovePaykToCustomCredit = True
            Case EnumActionLog.MoveTableToCustomCredit
                m_MoveTableToCustomCredit = True
            Case EnumActionLog.PosAcknowledged
                m_PosAcknowledged = True
                
        End Select
        Rst.MoveNext
    Wend
    
    Set Rst = Nothing
    Set ClsCnn = Nothing

End Sub
