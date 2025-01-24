VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mNo As Long
Private mStatus As Integer
Private mOwner As Integer
Private mCustomer As Long
Private mDiscountTotal As Double
Private mCarryFeeTotal As Double
Private mRecursive As Integer
Private mIncharge As Integer
Private mFacPayment As Boolean
Private mOrderType As Integer
Private mStationId As Integer
Private mServiceTotal As Double
Private mPackingTotal As Double
Private mBascoleNo As Integer
Private mTableNo As Integer
Private mUser As Integer
Private mDateInvoice As String
Private mDetailsString As String
Private msFactorReceived As String
Private mBalance As Boolean
Private mAccountYear As Integer
Private mNvcDescription As String
Private mlblPayFactorTotal As String
Private mRegDate As String
Private mTempAddress As String
Private mGuestNo As Long
Private mTaxTotal As Double
Private mDutyTotal As Double

Public Property Let No(ByVal vData As Long)
    mNo = vData
End Property

Public Property Get No() As Long
    No = mNo
End Property
Public Property Let Status(ByVal vData As Integer)
    mStatus = vData
End Property

Public Property Get Status() As Integer
    Status = mStatus
End Property
Public Property Let Owner(ByVal vData As Integer)
    mOwner = vData
End Property

Public Property Get Owner() As Integer
    Owner = mOwner
End Property
Public Property Let Customer(ByVal vData As Long)
    mCustomer = vData
End Property

Public Property Get Customer() As Long
    Customer = mCustomer
End Property
Public Property Let DiscountTotal(ByVal vData As Double)
    mDiscountTotal = vData
End Property

Public Property Get DiscountTotal() As Double
    DiscountTotal = mDiscountTotal
End Property
Public Property Let CarryFeeTotal(ByVal vData As Double)
    mCarryFeeTotal = vData
End Property

Public Property Get CarryFeeTotal() As Double
    CarryFeeTotal = mCarryFeeTotal
End Property
Public Property Let Recursive(ByVal vData As Integer)
    mRecursive = vData
End Property

Public Property Get Recursive() As Integer
    Recursive = mRecursive
End Property
Public Property Let Incharge(ByVal vData As Integer)
    mIncharge = vData
End Property

Public Property Get Incharge() As Integer
    Incharge = mIncharge
End Property
Public Property Let FacPayment(ByVal vData As Boolean)
    mFacPayment = vData
End Property

Public Property Get FacPayment() As Boolean
    FacPayment = mFacPayment
End Property
Public Property Let OrderType(ByVal vData As Integer)
    mOrderType = vData
End Property
Public Property Get OrderType() As Integer
    OrderType = mOrderType
End Property
Public Property Let StationId(ByVal vData As Integer)
    mStationId = vData
End Property

Public Property Get StationId() As Integer
    StationId = mStationId
End Property
Public Property Let ServiceTotal(ByVal vData As Double)
    mServiceTotal = vData
End Property

Public Property Get ServiceTotal() As Double
    ServiceTotal = mServiceTotal
End Property
Public Property Let PackingTotal(ByVal vData As Double)
    mPackingTotal = vData
End Property

Public Property Get PackingTotal() As Double
    PackingTotal = mPackingTotal
End Property
Public Property Let BascoleNo(ByVal vData As Integer)
    mBascoleNo = vData
End Property
Public Property Let TableNo(ByVal vData As Integer)
    mTableNo = vData
End Property

Public Property Get TableNo() As Integer
    TableNo = mTableNo
End Property

Public Property Get BascoleNo() As Integer
    BascoleNo = mBascoleNo
End Property

Public Property Let User(ByVal vData As Integer)
    mUser = vData
End Property

Public Property Get User() As Integer
    User = mUser
End Property
Public Property Let DateInvoice(ByVal vData As String)
    mDateInvoice = vData
End Property

Public Property Get DateInvoice() As String
    DateInvoice = mDateInvoice
End Property
Public Property Let DetailsString(ByVal vData As String)
    mDetailsString = vData
End Property

Public Property Get DetailsString() As String
    DetailsString = mDetailsString
End Property
Public Property Let sFactorReceived(ByVal vData As String)
    msFactorReceived = vData
End Property

Public Property Get sFactorReceived() As String
    sFactorReceived = msFactorReceived
End Property
Public Property Let Balance(ByVal vData As Boolean)
    mBalance = vData
End Property

Public Property Get Balance() As Boolean
    Balance = mBalance
End Property
Public Property Let AccountYear(ByVal vData As Integer)
    mAccountYear = vData
End Property

Public Property Get AccountYear() As Integer
    AccountYear = mAccountYear
End Property
Public Property Let NvcDescription(ByVal vData As String)
    mNvcDescription = vData
End Property

Public Property Get NvcDescription() As String
    NvcDescription = mNvcDescription
End Property
Public Property Let lblPayFactorTotal(ByVal vData As Double)
    mlblPayFactorTotal = vData
End Property

Public Property Get lblPayFactorTotal() As Double
    lblPayFactorTotal = mlblPayFactorTotal
End Property

Public Property Let RegDate(ByVal vData As String)
    mRegDate = vData
End Property
Public Property Get RegDate() As String
    RegDate = mRegDate
End Property
Public Property Let TempAddress(ByVal vData As String)
    mTempAddress = vData
End Property
Public Property Get TempAddress() As String
    TempAddress = mTempAddress
End Property
Public Property Let GuestNo(ByVal vData As Integer)
    mGuestNo = vData
End Property
Public Property Get GuestNo() As Integer
    GuestNo = mGuestNo
End Property

Public Property Let TaxTotal(ByVal vData As Double)
    mTaxTotal = vData
End Property

Public Property Get TaxTotal() As Double
    TaxTotal = mTaxTotal
End Property
Public Property Let DutyTotal(ByVal vData As Double)
    mDutyTotal = vData
End Property

Public Property Get DutyTotal() As Double
    DutyTotal = mDutyTotal
End Property

