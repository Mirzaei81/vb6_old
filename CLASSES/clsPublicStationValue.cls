VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicStationValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_Language As EnumLanguage

Private m_WinAscii As Integer
Private m_DrawerPort As Integer
Private m_DrawerModel As Integer
Private m_PartitionID As Integer
Private m_MaxAutoDiscount As Integer
Private m_CustDisplayPort As Integer
Private m_CustDisplayModel As Integer
Private m_DefaultCustSearch As Integer
Private m_ServePlaceDefault As Integer
Private m_BasculeModel As Integer
Private m_BasculePort As Integer
Private m_DeliveryBarcodeDefault As Integer
Private m_TableBarcodeDefault As Integer
Private m_ReprintDefault As Integer
Private m_KeyboardType As Integer
Private m_SearchType As Integer
Private m_PriceType As Integer
Private m_CustomerServeplace As Integer
Private m_DiscountDefault As Integer
Private m_SrarchInputDelayKeyboard As Integer
Private m_MaxRecordCount As Integer
Private m_FactorSortItems As Integer
Private m_EscapeInvoiceFactor As Integer
Private m_PurchaseInventoryDefault As Integer
Private m_StartUpFormDefault As Integer
Private m_Barcodelengh As Integer
Private m_MaxPrices As Integer
Private m_FromStoreFee As Integer
Private m_SaleStartDefault As Integer
Private m_ShowDigitNumber As Integer
Private m_CustomerRate As Integer
Private m_SearchOrderType As Integer
Private m_CallerIdSpace As Long
Private m_CountCustomerDailyBuy As Long
Private m_CountCustomerGood As Long
Private m_InvoiceRows As Integer
Private m_PurchaseRows As Integer
Private m_CycleStockNoDefault As Integer
Private m_NumberOfId As Integer
Private m_DiscoveryPort As Long
Private m_ResponsePort As Long
Private m_StartNumberCartReader As Integer
Private m_NumberOfCardReader As Integer
Private m_CityCode As String
Private m_TzWeightBarcodeId As String
Private m_OutPrice As Integer
Private m_TzNumericBarcodeId As String

Private m_TouchScreen As Boolean
Private m_CashDrawerPrn As Boolean
Private m_CustDisplayPrn As Boolean
Private m_DeliveryNoView As Boolean
Private m_RoundTwoNumber As Boolean
Private m_BasculeOn As Boolean
Private m_AutoDrawerOpen As Boolean
Private m_ChangeGoodPrint As Boolean
Private m_AlphabeticGoods As Boolean
Private m_CashClose As Boolean
Private m_AccCashClose As Boolean
Private m_DeletedGood As Boolean
Private m_SearchFichDefault As Boolean
Private m_CustomerOrderDefault As Boolean
Private m_CustomerSearchDefault As Boolean
Private m_CreditCalculate As Boolean
Private m_GoodSearchDefault As Boolean
Private m_MojodiControlDefault As Boolean
Private m_CommandView As Boolean
Private m_RowMojodiControl As Boolean
Private m_BarcodeChance As Boolean
Private m_RefreshFichNo As Boolean
Private m_DirectBascule As Boolean
Private m_PriceChance As String
Private m_StopOnEditFich As Boolean
Private m_MultiPrice As Boolean
Private m_UpdateBuyPrice As Boolean
Private m_UpdateSellPrice As Boolean
Private m_TrazooBarcode As Boolean
Private m_CodeFlag As Boolean
Private m_CircleMenu As Boolean
Private m_SoundAlarm As Boolean
Private m_TypeBascule As Integer
Private m_CashPayment As Boolean
Private m_ThreeSegmentSearch As Boolean
Private m_NumberOfUnitSale As Boolean
Private m_MenuViewAfterGood As Boolean
Private m_PayFactorView As Boolean
Private m_BarcodeAutoEscape As Boolean
Private m_PrintAfterTasvieh As Boolean
Private m_AutoBarcode As Boolean
Private m_AutoCallerId As Boolean
Private m_CustomerFarsi As Boolean
Private m_CustomerOnlinePrice As Boolean
Private m_CustomerAscii As Boolean
Private m_NoCurrentDay As Boolean
Private m_GoodPercentage As Boolean
Private m_TableControl As Boolean
'''Private m_Callwaiting As Boolean
Private m_FinalCheck As Boolean
Private m_AutoTip As Boolean
Private m_FichStatusBar As Boolean
Private m_TextIconViewH As Boolean
Private m_TextIconViewv As Boolean
Private m_ForceSeller As Boolean
Private m_InvoiceStatusDefault As Boolean
Private m_AlmLogFile As Boolean
Private m_NetworkCallerId As Boolean
Private m_CustomerFeeDataBase As Boolean
Private m_GoodLevelAutoDiscount As Boolean
Private m_ShiftRate As Boolean
Private m_SellerCaption As String
Private m_PrintAfterDeliver As Boolean

Private m_InpersonSalonPayment As Boolean
Private m_InpersonDeliveryPayment As Boolean
Private m_InpersonOutPayment As Boolean
Private m_InpersontablePayment As Boolean
Private m_InpersonSalonBalance As Boolean
Private m_InpersonDeliveryBalance As Boolean
Private m_InpersonOutBalance As Boolean
Private m_InpersontableBalance As Boolean

Private m_ByPhoneSalonPayment As Boolean
Private m_ByPhoneDeliveryPayment As Boolean
Private m_ByPhonetablePayment As Boolean
Private m_ByPhoneSalonBalance As Boolean
Private m_ByPhoneDeliveryBalance As Boolean
Private m_ByPhonetableBalance As Boolean
Private m_ReadDirectWeightTz1 As Boolean
Private m_TaxView As Boolean
Private m_NumberOfUnitBuy As Boolean
Private m_EditCompatibleSamar1 As Boolean
Private m_UndoRedoCompatibleSamar1 As Boolean
Private m_EscNotExit As Boolean
Private m_ViewTempAddress As Integer
Private m_CancelNotExit As Boolean
Private m_ResiveFormerFichCurrentDay As Boolean
Private m_SelectInventory As Boolean
Private m_AlphabetGoodSearch As Boolean
Private m_DeficitLog As Boolean
Private m_CallerId8Port As Boolean
Private m_SelectSeller As Boolean
Private m_CountCustomerShiftBuy As Long
Private m_FixRateChange As Boolean
Private m_AutoCashClose As Boolean
Private m_UpDateCarryFee As Boolean
Private m_ShowClock As Boolean

Private m_VoiceRecord As Boolean
Private m_treeViewMenu As Boolean
Private m_CallerIdAutoView As Boolean
Private m_CallerIdTest As Boolean
Private m_Pager As Boolean
Private m_AutoBackup As Boolean
Private m_PosPayment As Boolean
Private m_PassPhrase As String
Private m_PosApprovedText As String
Private m_PosPort As Integer
Private m_PosModel As Integer
Private m_TemporaryNo As Boolean
Private m_ReportHeadername As String
Private m_FastCustSave As Boolean
Private m_RepetitiveGood As Boolean
Private m_OtherPartition As Boolean
Private m_NotShowPrintNotice As Boolean
Private m_MultiInventory As Boolean
Private m_Frame_Printers As Boolean
Private m_HasOptionPrice As Boolean
Private m_ShowOption As Boolean
Private m_LableUsedGood As Boolean
Private m_LoyaltyCustomers As Boolean
Private m_LoyaltyAllCustomers As Boolean
Private m_StartCharacter As String
Private m_OneClickShow As Boolean
Private m_NoRowMenu As Integer
Private m_RfidReader As Boolean
Private m_RfidInterval As String
Private m_RfidLongBuzzer As Boolean
Private m_TelNetServerActive As Boolean
Private m_TelNetServerIP As String
Private m_TelNetServerPort As String
Private m_PrintAfterPayk As Boolean
Private m_PrintAfterOrder As Boolean
Private m_LabelPrint As Boolean
Private m_AryaSmsPanel As Boolean
Private m_ForceTax As Boolean
Private m_PersonIdCheck As Boolean
Private m_PersonIdRefreshTime As Integer
Private m_ListAutoLoad As Boolean
Private m_DeviceIP As String
Private m_DeviceID As Integer
Private m_ListFont As Integer
Private m_Device2IP As String
Private m_Device2ID As Integer

Public Property Get DefaultCustSearch() As Integer
    DefaultCustSearch = m_DefaultCustSearch
End Property
Public Property Let DefaultCustSearch(ByVal vData As Integer)
    m_DefaultCustSearch = vData
End Property
Public Property Get WinAscii() As Integer
    WinAscii = m_WinAscii
End Property
Public Property Get PartitionID() As Integer
    PartitionID = m_PartitionID
End Property
Public Property Let PartitionID(ByVal vData As Integer)
   m_PartitionID = vData
End Property

Public Property Get DeliveryNoView() As Boolean
    DeliveryNoView = m_DeliveryNoView
End Property
Public Property Let DeliveryNoView(ByVal vData As Boolean)
    m_DeliveryNoView = vData
End Property
Public Property Get AutoDrawerOpen() As Boolean
    AutoDrawerOpen = m_AutoDrawerOpen
End Property
Public Property Let AutoDrawerOpen(ByVal vData As Boolean)
    m_AutoDrawerOpen = vData
End Property
Public Property Get ChangeGoodPrint() As Boolean
    ChangeGoodPrint = m_ChangeGoodPrint
End Property
Public Property Let ChangeGoodPrint(ByVal vData As Boolean)
    m_ChangeGoodPrint = vData
End Property
Public Property Get AlphabeticGoods() As Boolean
    AlphabeticGoods = m_AlphabeticGoods
End Property
Public Property Let AlphabeticGoods(ByVal vData As Boolean)
    m_AlphabeticGoods = vData
End Property
Public Property Get TableControl() As Boolean
    TableControl = m_TableControl
End Property
Public Property Let TableControl(ByVal vData As Boolean)
    m_TableControl = vData
End Property
Public Property Get RoundTwoNumber() As Boolean
    RoundTwoNumber = m_RoundTwoNumber
End Property
Public Property Let RoundTwoNumber(ByVal vData As Boolean)
    m_RoundTwoNumber = vData
End Property
Public Property Get ServePlaceDefault() As Integer
    ServePlaceDefault = m_ServePlaceDefault
End Property
Public Property Let ServePlaceDefault(ByVal vData As Integer)
    m_ServePlaceDefault = vData
End Property
Public Property Get PurchaseInventoryDefault() As Integer
    PurchaseInventoryDefault = m_PurchaseInventoryDefault
End Property
Public Property Let PurchaseInventoryDefault(ByVal vData As Integer)
    m_PurchaseInventoryDefault = vData
End Property
Public Property Get MaxAutoDiscount() As Integer
    MaxAutoDiscount = m_MaxAutoDiscount
End Property
Public Property Let MaxAutoDiscount(ByVal vData As Integer)
    m_MaxAutoDiscount = vData
End Property
Public Property Get Language() As EnumLanguage
    Language = m_Language
End Property
Public Property Let Language(ByVal vData As EnumLanguage)
    m_Language = vData
End Property
Public Property Get TouchScreen() As Boolean
    TouchScreen = m_TouchScreen
End Property
Public Property Let TouchScreen(ByVal vData As Boolean)
    m_TouchScreen = vData
End Property
Public Property Get BasculeModel() As Integer
    BasculeModel = m_BasculeModel
End Property
Public Property Let BasculeModel(ByVal vData As Integer)
    m_BasculeModel = vData
End Property
Public Property Get BasculePort() As Integer
    BasculePort = m_BasculePort
End Property
Public Property Let BasculePort(ByVal vData As Integer)
    m_BasculePort = vData
End Property
Public Property Get DrawerModel() As Integer
    DrawerModel = m_DrawerModel
End Property
Public Property Let DrawerModel(ByVal vData As Integer)
    m_DrawerModel = vData
End Property
Public Property Get DrawerPort() As Integer
    DrawerPort = m_DrawerPort
End Property
Public Property Let DrawerPort(ByVal vData As Integer)
    m_DrawerPort = vData
End Property
Public Property Get CustDisplayModel() As Integer
    CustDisplayModel = m_CustDisplayModel
End Property
Public Property Let CustDisplayModel(ByVal vData As Integer)
    m_CustDisplayModel = vData
End Property
Public Property Get CustDisplayPort() As Integer
    CustDisplayPort = m_CustDisplayPort
End Property
Public Property Let CustDisplayPort(ByVal vData As Integer)
    m_CustDisplayPort = vData
End Property
Public Property Get CustDisplayPrn() As Boolean
    CustDisplayPrn = m_CustDisplayPrn
End Property
Public Property Let CustDisplayPrn(ByVal vData As Boolean)
    m_CustDisplayPrn = vData
End Property
Public Property Get CashDrawerPrn() As Boolean
    CashDrawerPrn = m_CashDrawerPrn
End Property
Public Property Let CashDrawerPrn(ByVal vData As Boolean)
    m_CashDrawerPrn = vData
End Property
Public Property Get DeliveryBarcodeDefault() As Integer
    DeliveryBarcodeDefault = m_DeliveryBarcodeDefault
End Property
Public Property Let DeliveryBarcodeDefault(ByVal vData As Integer)
    m_DeliveryBarcodeDefault = vData
End Property
Public Property Get TableBarcodeDefault() As Integer
    TableBarcodeDefault = m_TableBarcodeDefault
End Property
Public Property Let TableBarcodeDefault(ByVal vData As Integer)
    m_TableBarcodeDefault = vData
End Property
Public Property Get ReprintDefault() As Integer
    ReprintDefault = m_ReprintDefault
End Property
Public Property Let ReprintDefault(ByVal vData As Integer)
    m_ReprintDefault = vData
End Property
Public Property Get BasculeOn() As Boolean
    BasculeOn = m_BasculeOn
End Property
Public Property Let BasculeOn(ByVal vData As Boolean)
    m_BasculeOn = vData
End Property
Public Property Get KeyboardType() As Integer
    KeyboardType = m_KeyboardType
End Property
Public Property Let KeyboardType(ByVal vData As Integer)
    m_KeyboardType = vData
End Property
Public Property Get SearchType() As Integer
    SearchType = m_SearchType
End Property
Public Property Let SearchType(ByVal vData As Integer)
    m_SearchType = vData
End Property
Public Property Get PriceType() As Integer
    PriceType = m_PriceType
End Property
Public Property Let PriceType(ByVal vData As Integer)
    m_PriceType = vData
End Property
Public Property Get CashClose() As Boolean
    CashClose = m_CashClose
End Property
Public Property Let CashClose(ByVal vData As Boolean)
    m_CashClose = vData
End Property
Public Property Get AccCashClose() As Boolean
    AccCashClose = m_AccCashClose
End Property
Public Property Let AccCashClose(ByVal vData As Boolean)
    m_AccCashClose = vData
End Property
Public Property Get DeletedGood() As Boolean
    DeletedGood = m_DeletedGood
End Property
Public Property Let DeletedGood(ByVal vData As Boolean)
    m_DeletedGood = vData
End Property
Public Property Get SearchFichDefault() As Boolean
    SearchFichDefault = m_SearchFichDefault
End Property
Public Property Let SearchFichDefault(ByVal vData As Boolean)
    m_SearchFichDefault = vData
End Property
Public Property Get CustomerOrderDefault() As Boolean
    CustomerOrderDefault = m_CustomerOrderDefault
End Property
Public Property Let CustomerOrderDefault(ByVal vData As Boolean)
    m_CustomerOrderDefault = vData
End Property
Public Property Get CustomerServeplace() As Integer
    CustomerServeplace = m_CustomerServeplace
End Property
Public Property Let CustomerServeplace(ByVal vData As Integer)
    m_CustomerServeplace = vData
End Property
Public Property Get CustomerSearchDefault() As Boolean
    CustomerSearchDefault = m_CustomerSearchDefault
End Property
Public Property Let CustomerSearchDefault(ByVal vData As Boolean)
    m_CustomerSearchDefault = vData
End Property
Public Property Get CreditCalculate() As Boolean
    CreditCalculate = m_CreditCalculate
End Property
Public Property Let CreditCalculate(ByVal vData As Boolean)
    m_CreditCalculate = vData
End Property
Public Property Get GoodSearchDefault() As Boolean
    GoodSearchDefault = m_GoodSearchDefault
End Property
Public Property Let GoodSearchDefault(ByVal vData As Boolean)
    m_GoodSearchDefault = vData
End Property
Public Property Get DiscountDefault() As Integer
    DiscountDefault = m_DiscountDefault
End Property
Public Property Let DiscountDefault(ByVal vData As Integer)
    m_DiscountDefault = vData
End Property
Public Property Get StartUpFormDefault() As Integer
    StartUpFormDefault = m_StartUpFormDefault
End Property
Public Property Let StartUpFormDefault(ByVal vData As Integer)
    m_StartUpFormDefault = vData
End Property
Public Property Get SrarchInputDelayKeyboard() As Integer
    SrarchInputDelayKeyboard = m_SrarchInputDelayKeyboard
End Property
Public Property Let SrarchInputDelayKeyboard(ByVal vData As Integer)
    m_SrarchInputDelayKeyboard = vData
End Property
Public Property Get MaxRecordCount() As Integer
    MaxRecordCount = m_MaxRecordCount
End Property
Public Property Let MaxRecordCount(ByVal vData As Integer)
    m_MaxRecordCount = vData
End Property
Public Property Get FactorSortItems() As Integer
    FactorSortItems = m_FactorSortItems
End Property
Public Property Let FactorSortItems(ByVal vData As Integer)
    m_FactorSortItems = vData
End Property
Public Property Get EscapeInvoiceFactor() As Integer
    EscapeInvoiceFactor = m_EscapeInvoiceFactor
End Property

Public Property Get MojodiControlDefault() As Boolean
    MojodiControlDefault = m_MojodiControlDefault
End Property
Public Property Let MojodiControlDefault(ByVal vData As Boolean)
    m_MojodiControlDefault = vData
End Property
Public Property Get RowMojodiControl() As Boolean
    RowMojodiControl = m_RowMojodiControl
End Property
Public Property Let RowMojodiControl(ByVal vData As Boolean)
    m_RowMojodiControl = vData
End Property
Public Property Get CommandView() As Boolean
    CommandView = m_CommandView
End Property
Public Property Let CommandView(ByVal vData As Boolean)
    m_CommandView = vData
End Property
Public Property Get BarcodeLengh() As Integer
    BarcodeLengh = m_Barcodelengh
End Property
Public Property Let BarcodeLengh(ByVal vData As Integer)
    m_Barcodelengh = vData
End Property
Public Property Get BarcodeChance() As Boolean
    BarcodeChance = m_BarcodeChance
End Property
Public Property Let BarcodeChance(ByVal vData As Boolean)
    m_BarcodeChance = vData
End Property
Public Property Get PriceChance() As String
    PriceChance = m_PriceChance
End Property
Public Property Let PriceChance(ByVal vData As String)
    m_PriceChance = vData
End Property
Public Property Get RefreshFichNo() As Boolean
    RefreshFichNo = m_RefreshFichNo
End Property
Public Property Let RefreshFichNo(ByVal vData As Boolean)
    m_RefreshFichNo = vData
End Property
Public Property Get DirectBascule() As Boolean
    DirectBascule = m_DirectBascule
End Property
Public Property Let DirectBascule(ByVal vData As Boolean)
    m_DirectBascule = vData
End Property
Public Property Get StopOnEditFich() As Boolean
    StopOnEditFich = m_StopOnEditFich
End Property
Public Property Let StopOnEditFich(ByVal vData As Boolean)
    m_StopOnEditFich = vData
End Property
Public Property Get MultiPrice() As Boolean
    MultiPrice = m_MultiPrice
End Property
Public Property Let MultiPrice(ByVal vData As Boolean)
    m_MultiPrice = vData
End Property
Public Property Get UpdateBuyPrice() As Boolean
    UpdateBuyPrice = m_UpdateBuyPrice
End Property
Public Property Let UpdateBuyPrice(ByVal vData As Boolean)
    m_UpdateBuyPrice = vData
End Property
Public Property Get UpdateSellprice() As Boolean
    UpdateSellprice = m_UpdateSellPrice
End Property
Public Property Let UpdateSellprice(ByVal vData As Boolean)
    m_UpdateSellPrice = vData
End Property
Public Property Get TrazooBarcode() As Boolean
    TrazooBarcode = m_TrazooBarcode
End Property
Public Property Let TrazooBarcode(ByVal vData As Boolean)
    m_TrazooBarcode = vData
End Property
Public Property Get CodeFlag() As Boolean
    CodeFlag = m_CodeFlag
End Property
Public Property Let CodeFlag(ByVal vData As Boolean)
    m_CodeFlag = vData
End Property
Public Property Get CircleMenu() As Boolean
    CircleMenu = m_CircleMenu
End Property
Public Property Let CircleMenu(ByVal vData As Boolean)
    m_CircleMenu = vData
End Property
Public Property Get SoundAlarm() As Boolean
    SoundAlarm = m_SoundAlarm
End Property
Public Property Let SoundAlarm(ByVal vData As Boolean)
    m_SoundAlarm = vData
End Property
Public Property Get TypeBascule() As Integer
    TypeBascule = m_TypeBascule
End Property
Public Property Let TypeBascule(ByVal vData As Integer)
    m_TypeBascule = vData
End Property
Public Property Get CashPayment() As Boolean
    CashPayment = m_CashPayment
End Property
Public Property Let CashPayment(ByVal vData As Boolean)
    m_CashPayment = vData
End Property
'-------------------------------------
Public Property Get InpersonSalonPayment() As Boolean
    InpersonSalonPayment = m_InpersonSalonPayment
End Property
Public Property Let InpersonSalonPayment(ByVal vData As Boolean)
    m_InpersonSalonPayment = vData
End Property
Public Property Let InpersonDeliveryPayment(ByVal vData As Boolean)
    m_InpersonDeliveryPayment = vData
End Property
Public Property Get InpersonDeliveryPayment() As Boolean
    InpersonDeliveryPayment = m_InpersonDeliveryPayment
End Property
Public Property Let InpersonOutPayment(ByVal vData As Boolean)
    m_InpersonOutPayment = vData
End Property
Public Property Get InpersonOutPayment() As Boolean
    InpersonOutPayment = m_InpersonOutPayment
End Property
Public Property Get InpersonTablePayment() As Boolean
    InpersonTablePayment = m_InpersontablePayment
End Property
Public Property Let InpersonTablePayment(ByVal vData As Boolean)
    m_InpersontablePayment = vData
End Property
'-------------------------------------
'-------------------------------------
Public Property Get InpersonSalonBalance() As Boolean
    InpersonSalonBalance = m_InpersonSalonBalance
End Property
Public Property Let InpersonSalonBalance(ByVal vData As Boolean)
    m_InpersonSalonBalance = vData
End Property
Public Property Let InpersonDeliveryBalance(ByVal vData As Boolean)
    m_InpersonDeliveryBalance = vData
End Property
Public Property Get InpersonDeliveryBalance() As Boolean
    InpersonDeliveryBalance = m_InpersonDeliveryBalance
End Property
Public Property Let InpersonOutBalance(ByVal vData As Boolean)
    m_InpersonOutBalance = vData
End Property
Public Property Get InpersonOutBalance() As Boolean
    InpersonOutBalance = m_InpersonOutBalance
End Property
Public Property Get InpersonTableBalance() As Boolean
    InpersonTableBalance = m_InpersontableBalance
End Property
Public Property Let InpersonTableBalance(ByVal vData As Boolean)
    m_InpersontableBalance = vData
End Property
'-------------------------------------
'-------------------------------------
Public Property Get ByPhoneSalonPayment() As Boolean
    ByPhoneSalonPayment = m_ByPhoneSalonPayment
End Property
Public Property Let ByPhoneSalonPayment(ByVal vData As Boolean)
    m_ByPhoneSalonPayment = vData
End Property
Public Property Get ByPhoneDeliveryPayment() As Boolean
    ByPhoneDeliveryPayment = m_ByPhoneDeliveryPayment
End Property
Public Property Let ByPhoneDeliveryPayment(ByVal vData As Boolean)
    m_ByPhoneDeliveryPayment = vData
End Property
Public Property Get ByPhoneTablePayment() As Boolean
    ByPhoneTablePayment = m_ByPhonetablePayment
End Property
Public Property Let ByPhoneTablePayment(ByVal vData As Boolean)
    m_ByPhonetablePayment = vData
End Property
'-------------------------------------
'-------------------------------------
Public Property Get ByPhoneSalonBalance() As Boolean
    ByPhoneSalonBalance = m_ByPhoneSalonBalance
End Property
Public Property Let ByPhoneSalonBalance(ByVal vData As Boolean)
    m_ByPhoneSalonBalance = vData
End Property
Public Property Get ByPhoneDeliveryBalance() As Boolean
    ByPhoneDeliveryBalance = m_ByPhoneDeliveryBalance
End Property
Public Property Let ByPhoneDeliveryBalance(ByVal vData As Boolean)
    m_ByPhoneDeliveryBalance = vData
End Property
Public Property Get ByPhoneTableBalance() As Boolean
    ByPhoneTableBalance = m_ByPhonetableBalance
End Property
Public Property Let ByPhoneTableBalance(ByVal vData As Boolean)
    m_ByPhonetableBalance = vData
End Property
Public Property Get ThreeSegmentSearch() As Boolean
    ThreeSegmentSearch = m_ThreeSegmentSearch
End Property
Public Property Let ThreeSegmentSearch(ByVal vData As Boolean)
    m_ThreeSegmentSearch = vData
End Property
Public Property Get MaxPrices() As Integer
    MaxPrices = m_MaxPrices
End Property
Public Property Let MaxPrices(ByVal vData As Integer)
    m_MaxPrices = vData
End Property
Public Property Get FromStoreFee() As Integer
    FromStoreFee = m_FromStoreFee
End Property
Public Property Let FromStoreFee(ByVal vData As Integer)
    m_FromStoreFee = vData
End Property

Public Property Get NumberOfUnitSale() As Boolean
    NumberOfUnitSale = m_NumberOfUnitSale
End Property
Public Property Let NumberOfUnitSale(ByVal vData As Boolean)
    m_NumberOfUnitSale = vData
End Property
Public Property Get MenuViewAfterGood() As Boolean
    MenuViewAfterGood = m_MenuViewAfterGood
End Property
Public Property Let MenuViewAfterGood(ByVal vData As Boolean)
    m_MenuViewAfterGood = vData
End Property
Public Property Get PayFactorView() As Boolean
    PayFactorView = m_PayFactorView
End Property
Public Property Let PayFactorView(ByVal vData As Boolean)
    m_PayFactorView = vData
End Property
Public Property Get BarcodeAutoEscape() As Boolean
    BarcodeAutoEscape = m_BarcodeAutoEscape
End Property
Public Property Let BarcodeAutoEscape(ByVal vData As Boolean)
    m_BarcodeAutoEscape = vData
End Property
Public Property Get PrintAfterTasvieh() As Boolean
    PrintAfterTasvieh = m_PrintAfterTasvieh
End Property
Public Property Let PrintAfterTasvieh(ByVal vData As Boolean)
    m_PrintAfterTasvieh = vData
End Property
Public Property Get AutoBarcode() As Boolean
    AutoBarcode = m_AutoBarcode
End Property
Public Property Let AutoBarcode(ByVal vData As Boolean)
    m_AutoBarcode = vData
End Property
Public Property Get AutoCallerId() As Boolean
    AutoCallerId = m_AutoCallerId
End Property
Public Property Let AutoCallerId(ByVal vData As Boolean)
    m_AutoCallerId = vData
End Property
Public Property Get CustomerOnlinePrice() As Boolean
    CustomerOnlinePrice = m_CustomerOnlinePrice
End Property
Public Property Let CustomerOnlinePrice(ByVal vData As Boolean)
    m_CustomerOnlinePrice = vData
End Property
Public Property Get CustomerFarsi() As Boolean
    CustomerFarsi = m_CustomerFarsi
End Property
Public Property Let CustomerFarsi(ByVal vData As Boolean)
    m_CustomerFarsi = vData
End Property

Public Property Get CustomerAscii() As Boolean
    CustomerAscii = m_CustomerAscii
End Property
Public Property Let CustomerAscii(ByVal vData As Boolean)
    m_CustomerAscii = vData
End Property
Public Property Get NoCurrentDay() As Boolean
    NoCurrentDay = m_NoCurrentDay
End Property
Public Property Let NoCurrentDay(ByVal vData As Boolean)
    m_NoCurrentDay = vData
End Property
Public Property Get SaleStartDefault() As Integer
    SaleStartDefault = m_SaleStartDefault
End Property
Public Property Let SaleStartDefault(ByVal vData As Integer)
    m_SaleStartDefault = vData
End Property
Public Property Get ShowDigitNumber() As Integer
    ShowDigitNumber = m_ShowDigitNumber
End Property
Public Property Let ShowDigitNumber(ByVal vData As Integer)
    m_ShowDigitNumber = vData
End Property
Public Property Get CustomerRate() As Integer
    CustomerRate = m_CustomerRate
End Property
Public Property Let CustomerRate(ByVal vData As Integer)
    m_CustomerRate = vData
End Property
Public Property Get GoodPercentage() As Boolean
    GoodPercentage = m_GoodPercentage
End Property
Public Property Let GoodPercentage(ByVal vData As Boolean)
    m_GoodPercentage = vData
End Property
Public Property Get SearchOrderType() As Integer
    SearchOrderType = m_SearchOrderType
End Property
Public Property Let SearchOrderType(ByVal vData As Integer)
    m_SearchOrderType = vData
End Property
Public Property Get CallerIdSpace() As Long
    CallerIdSpace = m_CallerIdSpace
End Property
Public Property Let CallerIdSpace(ByVal vData As Long)
    m_CallerIdSpace = vData
End Property
'''Public Property Get Callwaiting() As Boolean
'''    Callwaiting = m_Callwaiting
'''End Property
'''Public Property Let Callwaiting(ByVal vData As Boolean)
'''    m_Callwaiting = vData
'''End Property
Public Property Get CountCustomerDailyBuy() As Long
    CountCustomerDailyBuy = m_CountCustomerDailyBuy
End Property
Public Property Let CountCustomerDailyBuy(ByVal vData As Long)
    m_CountCustomerDailyBuy = vData
End Property
Public Property Get CountCustomerGood() As Long
    CountCustomerGood = m_CountCustomerGood
End Property
Public Property Let CountCustomerGood(ByVal vData As Long)
    m_CountCustomerGood = vData
End Property
Public Property Get InvoiceRows() As Integer
    InvoiceRows = m_InvoiceRows
End Property
Public Property Let InvoiceRows(ByVal vData As Integer)
    m_InvoiceRows = vData
End Property
Public Property Get PurchaseRows() As Integer
    PurchaseRows = m_PurchaseRows
End Property
Public Property Let PurchaseRows(ByVal vData As Integer)
    m_PurchaseRows = vData
End Property
Public Property Get FinalCheck() As Boolean
    FinalCheck = m_FinalCheck
End Property
Public Property Let FinalCheck(ByVal vData As Boolean)
    m_FinalCheck = vData
End Property
Public Property Get AutoTip() As Boolean
    AutoTip = m_AutoTip
End Property
Public Property Let AutoTip(ByVal vData As Boolean)
    m_AutoTip = vData
End Property
Public Property Get FichStatusBar() As Boolean
    FichStatusBar = m_FichStatusBar
End Property
Public Property Let FichStatusBar(ByVal vData As Boolean)
    m_FichStatusBar = vData
End Property
Public Property Get CycleStockNoDefault() As Integer
    CycleStockNoDefault = m_CycleStockNoDefault
End Property
Public Property Let CycleStockNoDefault(ByVal vData As Integer)
    m_CycleStockNoDefault = vData
End Property
Public Property Get TextIconViewH() As Boolean
    TextIconViewH = m_TextIconViewH
End Property
Public Property Let TextIconViewH(ByVal vData As Boolean)
    m_TextIconViewH = vData
End Property
Public Property Get TextIconViewv() As Boolean
    TextIconViewv = m_TextIconViewv
End Property
Public Property Let TextIconViewv(ByVal vData As Boolean)
    m_TextIconViewv = vData
End Property
Public Property Get ForceSeller() As Boolean
    ForceSeller = m_ForceSeller
End Property
Public Property Let ForceSeller(ByVal vData As Boolean)
    m_ForceSeller = vData
End Property
Public Property Get InvoiceStatusDefault() As Boolean
    InvoiceStatusDefault = m_InvoiceStatusDefault
End Property
Public Property Let InvoiceStatusDefault(ByVal vData As Boolean)
    m_InvoiceStatusDefault = vData
End Property
Public Property Get NumberOfId() As Integer
    NumberOfId = m_NumberOfId
End Property
Public Property Let NumberOfId(ByVal vData As Integer)
   m_NumberOfId = vData
End Property
Public Property Get DiscoveryPort() As Integer
    DiscoveryPort = m_DiscoveryPort
End Property
Public Property Let DiscoveryPort(ByVal vData As Integer)
   m_DiscoveryPort = vData
End Property

Public Property Get ResponsePort() As Integer
    ResponsePort = m_ResponsePort
End Property
Public Property Let ResponsePort(ByVal vData As Integer)
   m_ResponsePort = vData
End Property
Public Property Get CityCode() As String
    CityCode = m_CityCode
End Property
Public Property Let CityCode(ByVal vData As String)
   m_CityCode = vData
End Property
Public Property Get AlmLogFile() As Boolean
    AlmLogFile = m_AlmLogFile
End Property
Public Property Let AlmLogFile(ByVal vData As Boolean)
    m_AlmLogFile = vData
End Property
Public Property Get NetworkCallerId() As Boolean
    NetworkCallerId = m_NetworkCallerId
End Property
Public Property Let NetworkCallerId(ByVal vData As Boolean)
    m_NetworkCallerId = vData
End Property

Public Property Get TzWeightBarcodeId() As String
    TzWeightBarcodeId = m_TzWeightBarcodeId
End Property
Public Property Let TzWeightBarcodeId(ByVal vData As String)
   m_TzWeightBarcodeId = vData
End Property
Public Property Get CustomerFeeDataBase() As Boolean
    CustomerFeeDataBase = m_CustomerFeeDataBase
End Property
Public Property Let CustomerFeeDataBase(ByVal vData As Boolean)
    m_CustomerFeeDataBase = vData
End Property
Public Property Get GoodLevelAutoDiscount() As Boolean
    GoodLevelAutoDiscount = m_GoodLevelAutoDiscount
End Property
Public Property Let GoodLevelAutoDiscount(ByVal vData As Boolean)
    m_GoodLevelAutoDiscount = vData
End Property
Public Property Get StartNumberCartReader() As Integer
    StartNumberCartReader = m_StartNumberCartReader
End Property
Public Property Let StartNumberCartReader(ByVal vData As Integer)
   m_StartNumberCartReader = vData
End Property
Public Property Get NumberOfCardReader() As Integer
    NumberOfCardReader = m_NumberOfCardReader
End Property
Public Property Let NumberOfCardReader(ByVal vData As Integer)
   m_NumberOfCardReader = vData
End Property
Public Property Get ReadDirectWeightTz1() As Boolean
    ReadDirectWeightTz1 = m_ReadDirectWeightTz1
End Property
Public Property Let ReadDirectWeightTz1(ByVal vData As Boolean)
    m_ReadDirectWeightTz1 = vData
End Property
Public Property Get TaxView() As Boolean
    TaxView = m_TaxView
End Property
Public Property Let TaxView(ByVal vData As Boolean)
    m_TaxView = vData
End Property
Public Property Get OutPrice() As Integer
    OutPrice = m_OutPrice
End Property
Public Property Let OutPrice(ByVal vData As Integer)
    m_OutPrice = vData
End Property
Public Property Get TzNumericBarcodeId() As String
    TzNumericBarcodeId = m_TzNumericBarcodeId
End Property
Public Property Let TzNumericBarcodeId(ByVal vData As String)
   m_TzNumericBarcodeId = vData
End Property
Public Property Get NumberOfUnitBuy() As Boolean
    NumberOfUnitBuy = m_NumberOfUnitBuy
End Property
Public Property Let NumberOfUnitBuy(ByVal vData As Boolean)
    m_NumberOfUnitBuy = vData
End Property
Public Property Get EditCompatibleSamar1() As Boolean
    EditCompatibleSamar1 = m_EditCompatibleSamar1
End Property
Public Property Let EditCompatibleSamar1(ByVal vData As Boolean)
    m_EditCompatibleSamar1 = vData
End Property
Public Property Get UndoRedoCompatibleSamar1() As Boolean
    UndoRedoCompatibleSamar1 = m_UndoRedoCompatibleSamar1
End Property
Public Property Let UndoRedoCompatibleSamar1(ByVal vData As Boolean)
    m_UndoRedoCompatibleSamar1 = vData
End Property
Public Property Get EscNotExit() As Boolean
    EscNotExit = m_EscNotExit
End Property
Public Property Let EscNotExit(ByVal vData As Boolean)
    m_EscNotExit = vData
End Property
Public Property Get ViewTempAddress() As Integer
    ViewTempAddress = m_ViewTempAddress
End Property
Public Property Let ViewTempAddress(ByVal vData As Integer)
    m_ViewTempAddress = vData
End Property
Public Property Get CancelNotExit() As Boolean
    CancelNotExit = m_CancelNotExit
End Property
Public Property Let CancelNotExit(ByVal vData As Boolean)
    m_CancelNotExit = vData
End Property
Public Property Get ResiveFormerFichCurrentDay() As Boolean
    ResiveFormerFichCurrentDay = m_ResiveFormerFichCurrentDay
End Property
Public Property Let ResiveFormerFichCurrentDay(ByVal vData As Boolean)
    m_ResiveFormerFichCurrentDay = vData
End Property
Public Property Get SelectInventory() As Boolean
    SelectInventory = m_SelectInventory
End Property
Public Property Let SelectInventory(ByVal vData As Boolean)
    m_SelectInventory = vData
End Property
Public Property Get AlphabetGoodSearch() As Boolean
    AlphabetGoodSearch = m_AlphabetGoodSearch
End Property
Public Property Let AlphabetGoodSearch(ByVal vData As Boolean)
    m_AlphabetGoodSearch = vData
End Property
Public Property Get DeficitLog() As Boolean
    DeficitLog = m_DeficitLog
End Property
Public Property Let DeficitLog(ByVal vData As Boolean)
    m_DeficitLog = vData
End Property
Public Property Get CallerId8Port() As Boolean
    CallerId8Port = m_CallerId8Port
End Property
Public Property Let CallerId8Port(ByVal vData As Boolean)
    m_CallerId8Port = vData
End Property
Public Property Get ShiftRate() As Boolean
    ShiftRate = m_ShiftRate
End Property
Public Property Let ShiftRate(ByVal vData As Boolean)
    m_ShiftRate = vData
End Property
Public Property Get SellerCaption() As String
    SellerCaption = m_SellerCaption
End Property
Public Property Let SellerCaption(ByVal vData As String)
    m_SellerCaption = vData
End Property
Public Property Get SelectSeller() As Boolean
    SelectSeller = m_SelectSeller
End Property
Public Property Let SelectSeller(ByVal vData As Boolean)
    m_SelectSeller = vData
End Property
Public Property Get CountCustomerShiftBuy() As Long
    CountCustomerShiftBuy = m_CountCustomerShiftBuy
End Property
Public Property Let CountCustomerShiftBuy(ByVal vData As Long)
    m_CountCustomerShiftBuy = vData
End Property
Public Property Get PrintAfterDeliver() As Boolean
    PrintAfterDeliver = m_PrintAfterDeliver
End Property
Public Property Let PrintAfterDeliver(ByVal vData As Boolean)
    m_PrintAfterDeliver = vData
End Property
Public Property Get FixRateChange() As Boolean
    FixRateChange = m_FixRateChange
End Property
Public Property Let FixRateChange(ByVal vData As Boolean)
    m_FixRateChange = vData
End Property
Public Property Get AutoCashClose() As Boolean
    AutoCashClose = m_AutoCashClose
End Property
Public Property Let AutoCashClose(ByVal vData As Boolean)
    m_AutoCashClose = vData
End Property
Public Property Get UpDateCarryFee() As Boolean
    UpDateCarryFee = m_UpDateCarryFee
End Property
Public Property Let UpDateCarryFee(ByVal vData As Boolean)
    m_UpDateCarryFee = vData
End Property
Public Property Get ShowClock() As Boolean
    ShowClock = m_ShowClock
End Property
Public Property Let ShowClock(ByVal vData As Boolean)
    m_ShowClock = vData
End Property
Public Property Get VoiceRecord() As Boolean
    VoiceRecord = m_VoiceRecord
End Property
Public Property Let VoiceRecord(ByVal vData As Boolean)
    m_VoiceRecord = vData
End Property
Public Property Get TreeViewMenu() As Boolean
    TreeViewMenu = m_treeViewMenu
End Property
Public Property Let TreeViewMenu(ByVal vData As Boolean)
    m_treeViewMenu = vData
End Property
Public Property Get CallerIdAutoView() As Boolean
    CallerIdAutoView = m_CallerIdAutoView
End Property
Public Property Let CallerIdAutoView(ByVal vData As Boolean)
    m_CallerIdAutoView = vData
End Property
Public Property Get CallerIdTest() As Boolean
    CallerIdTest = m_CallerIdTest
End Property
Public Property Let CallerIdTest(ByVal vData As Boolean)
    m_CallerIdTest = vData
End Property
Public Property Get Pager() As Boolean
    Pager = m_Pager
End Property
Public Property Let Pager(ByVal vData As Boolean)
    m_Pager = vData
End Property
Public Property Get AutoBackup() As Boolean
    AutoBackup = m_AutoBackup
End Property
Public Property Let AutoBackup(ByVal vData As Boolean)
    m_AutoBackup = vData
End Property
Public Property Get PosModel() As Integer
    PosModel = m_PosModel
End Property
Public Property Let PosModel(ByVal vData As Integer)
    m_PosModel = vData
End Property
Public Property Get PosPort() As Integer
    PosPort = m_PosPort
End Property
Public Property Let PosPort(ByVal vData As Integer)
    m_PosPort = vData
End Property
Public Property Get PosPayment() As Boolean
    PosPayment = m_PosPayment
End Property
Public Property Let PosPayment(ByVal vData As Boolean)
    m_PosPayment = vData
End Property
Public Property Get PassPhrase() As String
    PassPhrase = m_PassPhrase
End Property
Public Property Let PassPhrase(ByVal vData As String)
    m_PassPhrase = vData
End Property
Public Property Get PosApprovedText() As String
    PosApprovedText = m_PosApprovedText
End Property
Public Property Let PosApprovedText(ByVal vData As String)
    m_PosApprovedText = vData
End Property
Public Property Get TemporaryNo() As Boolean
    TemporaryNo = m_TemporaryNo
End Property
Public Property Let TemporaryNo(ByVal vData As Boolean)
    m_TemporaryNo = vData
End Property
Public Property Get ReportHeadername() As String
    ReportHeadername = m_ReportHeadername
End Property
Public Property Let ReportHeadername(ByVal vData As String)
    m_ReportHeadername = vData
End Property
Public Property Get FastCustSave() As Boolean
    FastCustSave = m_FastCustSave
End Property
Public Property Let FastCustSave(ByVal vData As Boolean)
    m_FastCustSave = vData
End Property
Public Property Get RepetitiveGood() As Boolean
    RepetitiveGood = m_RepetitiveGood
End Property
Public Property Let RepetitiveGood(ByVal vData As Boolean)
    m_RepetitiveGood = vData
End Property
Public Property Get OtherPartition() As Boolean
    OtherPartition = m_OtherPartition
End Property
Public Property Let OtherPartition(ByVal vData As Boolean)
    m_OtherPartition = vData
End Property
Public Property Get NotShowPrintNotice() As Boolean
    NotShowPrintNotice = m_NotShowPrintNotice
End Property
Public Property Let NotShowPrintNotice(ByVal vData As Boolean)
    m_NotShowPrintNotice = vData
End Property
Public Property Get MultiInventory() As Boolean
    MultiInventory = m_MultiInventory
End Property
Public Property Let MultiInventory(ByVal vData As Boolean)
    m_MultiInventory = vData
End Property
Public Property Get Frame_Printers() As Boolean
    Frame_Printers = m_Frame_Printers
End Property
Public Property Let Frame_Printers(ByVal vData As Boolean)
    m_Frame_Printers = vData
End Property
Public Property Get HasOptionPrice() As Boolean
    HasOptionPrice = m_HasOptionPrice
End Property
Public Property Let HasOptionPrice(ByVal vData As Boolean)
    m_HasOptionPrice = vData
End Property
Public Property Get ShowOption() As Boolean
    ShowOption = m_ShowOption
End Property
Public Property Let ShowOption(ByVal vData As Boolean)
    m_ShowOption = vData
End Property
Public Property Get LableUsedGood() As Boolean
    LableUsedGood = m_LableUsedGood
End Property
Public Property Let LableUsedGood(ByVal vData As Boolean)
    m_LableUsedGood = vData
End Property
Public Property Get LoyaltyCustomers() As Boolean
    LoyaltyCustomers = m_LoyaltyCustomers
End Property
Public Property Let LoyaltyCustomers(ByVal vData As Boolean)
    m_LoyaltyCustomers = vData
End Property
Public Property Get LoyaltyAllCustomers() As Boolean
    LoyaltyAllCustomers = m_LoyaltyAllCustomers
End Property
Public Property Let LoyaltyAllCustomers(ByVal vData As Boolean)
    m_LoyaltyAllCustomers = vData
End Property
Public Property Get StartCharacter() As String
    StartCharacter = m_StartCharacter
End Property
Public Property Let StartCharacter(ByVal vData As String)
    m_StartCharacter = vData
End Property
Public Property Get OneClickShow() As Boolean
    OneClickShow = m_OneClickShow
End Property
Public Property Let OneClickShow(ByVal vData As Boolean)
    m_OneClickShow = vData
End Property
Public Property Get NoRowMenu() As Integer
    NoRowMenu = m_NoRowMenu
End Property
Public Property Let NoRowMenu(ByVal vData As Integer)
    m_NoRowMenu = vData
End Property
Public Property Get RfidReader() As Boolean
    RfidReader = m_RfidReader
End Property
Public Property Let RfidReader(ByVal vData As Boolean)
    m_RfidReader = vData
End Property
Public Property Get RfidInterval() As String
    RfidInterval = m_RfidInterval
End Property
Public Property Let RfidInterval(ByVal vData As String)
    m_RfidInterval = vData
End Property
Public Property Get RfidLongBuzzer() As Boolean
    RfidLongBuzzer = m_RfidLongBuzzer
End Property
Public Property Let RfidLongBuzzer(ByVal vData As Boolean)
    m_RfidLongBuzzer = vData
End Property
Public Property Get TelNetServerActive() As Boolean
    TelNetServerActive = m_TelNetServerActive
End Property
Public Property Let TelNetServerActive(ByVal vData As Boolean)
    m_TelNetServerActive = vData
End Property
Public Property Get TelNetServerIP() As String
    TelNetServerIP = m_TelNetServerIP
End Property
Public Property Let TelNetServerIP(ByVal vData As String)
    m_TelNetServerIP = vData
End Property
Public Property Get TelNetServerPort() As String
    TelNetServerPort = m_TelNetServerPort
End Property
Public Property Let TelNetServerPort(ByVal vData As String)
    m_TelNetServerPort = vData
End Property
Public Property Get PrintAfterPayk() As Boolean
    PrintAfterPayk = m_PrintAfterPayk
End Property
Public Property Let PrintAfterPayk(ByVal vData As Boolean)
    m_PrintAfterPayk = vData
End Property
Public Property Get PrintAfterOrder() As Boolean
    PrintAfterOrder = m_PrintAfterOrder
End Property
Public Property Let PrintAfterOrder(ByVal vData As Boolean)
    m_PrintAfterOrder = vData
End Property
Public Property Get LabelPrint() As Boolean
    LabelPrint = m_LabelPrint
End Property
Public Property Let LabelPrint(ByVal vData As Boolean)
    m_LabelPrint = vData
End Property
Public Property Get AryaSmsPanel() As Boolean
    AryaSmsPanel = m_AryaSmsPanel
End Property
Public Property Let AryaSmsPanel(ByVal vData As Boolean)
    m_AryaSmsPanel = vData
End Property
Public Property Get ForceTax() As Boolean
    ForceTax = m_ForceTax
End Property
Public Property Let ForceTax(ByVal vData As Boolean)
    m_ForceTax = vData
End Property
Public Property Get PersonIdCheck() As Boolean
    PersonIdCheck = m_PersonIdCheck
End Property
Public Property Let PersonIdCheck(ByVal vData As Boolean)
    m_PersonIdCheck = vData
End Property
Public Property Get PersonIdRefreshTime() As Integer
    PersonIdRefreshTime = m_PersonIdRefreshTime
End Property
Public Property Let PersonIdRefreshTime(ByVal vData As Integer)
   m_PersonIdRefreshTime = vData
End Property
Public Property Get ListAutoLoad() As Boolean
    ListAutoLoad = m_ListAutoLoad
End Property
Public Property Let ListAutoLoad(ByVal vData As Boolean)
    m_ListAutoLoad = vData
End Property
Public Property Get DeviceIP() As String
    DeviceIP = m_DeviceIP
End Property
Public Property Let DeviceIP(ByVal vData As String)
   m_DeviceIP = vData
End Property
Public Property Get DeviceID() As Integer
    DeviceID = m_DeviceID
End Property
Public Property Let DeviceID(ByVal vData As Integer)
   m_DeviceID = vData
End Property
Public Property Get ListFont() As Integer
    ListFont = m_ListFont
End Property
Public Property Let ListFont(ByVal vData As Integer)
   m_ListFont = vData
End Property
Public Property Get Device2IP() As String
    Device2IP = m_Device2IP
End Property
Public Property Let Device2IP(ByVal vData As String)
   m_Device2IP = vData
End Property
Public Property Get Device2Id() As Integer
    Device2Id = m_Device2ID
End Property
Public Property Let Device2Id(ByVal vData As Integer)
   m_Device2ID = vData
End Property

Public Sub Class_Initialize()
'On Error Resume Next
    
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
    
    Dim IsFileExist As Boolean
    
    If StationSettingFile = "" Then End    'Only  For  Make Exe File
    Set tempstring = filetemp.OpenTextFile(StationSettingFile, ForReading, False, TristateFalse)
    
    Do While tempstring.AtEndOfLine = False
       Str = tempstring.ReadLine
       LenghStr = InStr(1, Str, "=", vbTextCompare)
       
       If InStr(1, Str, "MaxAutoDiscount", vbTextCompare) Then
          m_MaxAutoDiscount = Val(Trim(Trim(Mid(Str, LenghStr + 1))))
       
       ElseIf InStr(1, Str, "Language", vbTextCompare) Then
          m_Language = Trim(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "TouchScreen", vbTextCompare) Then
          m_TouchScreen = Trim(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "ServePlaceDefault", vbTextCompare) Then
          m_ServePlaceDefault = Val(Trim(Trim(Mid(Str, LenghStr + 1))))
       
       ElseIf InStr(1, Str, "PurchaseInventoryDefault", vbTextCompare) Then
          m_PurchaseInventoryDefault = Val(Trim(Trim(Mid(Str, LenghStr + 1))))
       
       ElseIf InStr(1, Str, "PartitionID", vbTextCompare) Then
          m_PartitionID = CInt(Trim(Trim(Mid(Str, LenghStr + 1))))
       
       ElseIf InStr(1, Str, "WinAscii", vbTextCompare) Then
          m_WinAscii = CInt(Trim(Trim(Mid(Str, LenghStr + 1))))
       
       ElseIf InStr(1, Str, "DefaultCustSearch", vbTextCompare) Then
          m_DefaultCustSearch = CInt(Trim(Trim(Mid(Str, LenghStr + 1))))
       
       ElseIf InStr(1, Str, "DeliveryBarcodeDefault", vbTextCompare) Then
          m_DeliveryBarcodeDefault = CInt(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "TableBarcodeDefault", vbTextCompare) Then
          m_TableBarcodeDefault = CInt(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "ReprintDefault", vbTextCompare) Then
          m_ReprintDefault = CInt(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "DeliveryNoView", vbTextCompare) Then
          m_DeliveryNoView = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "AutoDrawerOpen", vbTextCompare) Then
          m_AutoDrawerOpen = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "ChangeGoodPrint", vbTextCompare) Then
          m_ChangeGoodPrint = Trim(Mid(Str, LenghStr + 1))
            
       ElseIf InStr(1, Str, "AlphabeticGoods", vbTextCompare) Then
          m_AlphabeticGoods = Trim(Mid(Str, LenghStr + 1))
      
       ElseIf InStr(1, Str, "TableControl", vbTextCompare) Then
          m_TableControl = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "RoundTwoNumber", vbTextCompare) Then
          m_RoundTwoNumber = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "KeyboardType", vbTextCompare) Then
          m_KeyboardType = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "SearchType", vbTextCompare) Then
          m_SearchType = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "PriceType", vbTextCompare) Then
          m_PriceType = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "DeletedGood", vbTextCompare) Then
          m_DeletedGood = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "SearchFichDefault", vbTextCompare) Then
          m_SearchFichDefault = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "CustomerOrderDefault", vbTextCompare) Then
          m_CustomerOrderDefault = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "CustomerServeplace", vbTextCompare) Then
          m_CustomerServeplace = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "CustomerSearchDefault", vbTextCompare) Then
          m_CustomerSearchDefault = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "CreditCalculate", vbTextCompare) Then
          m_CreditCalculate = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "GoodSearchDefault", vbTextCompare) Then
          m_GoodSearchDefault = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "DiscountDefault", vbTextCompare) Then
          m_DiscountDefault = CInt(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "SrarchInputDelayKeyboard", vbTextCompare) Then
          m_SrarchInputDelayKeyboard = CInt(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "MaxRecordCount", vbTextCompare) Then
          m_MaxRecordCount = CInt(Trim(Mid(Str, LenghStr + 1)))
       
      ElseIf InStr(1, Str, "FactorSortItems", vbTextCompare) Then
          m_FactorSortItems = CInt(Trim(Mid(Str, LenghStr + 1)))
       
      ElseIf InStr(1, Str, "MojodiControlDefault", vbTextCompare) Then
          m_MojodiControlDefault = Trim(Mid(Str, LenghStr + 1))
       
      ElseIf InStr(1, Str, "RowMojodiControl", vbTextCompare) Then
          m_RowMojodiControl = Trim(Mid(Str, LenghStr + 1))
       
      ElseIf InStr(1, Str, "CommandView", vbTextCompare) Then
          m_CommandView = Trim(Mid(Str, LenghStr + 1))
       
      ElseIf InStr(1, Str, "EscapeInvoiceFactor", vbTextCompare) Then
          m_EscapeInvoiceFactor = Trim(Mid(Str, LenghStr + 1))
      
      ElseIf InStr(1, Str, "StartUpFormDefault", vbTextCompare) Then
          m_StartUpFormDefault = CInt(Trim(Mid(Str, LenghStr + 1)))
       
      ElseIf InStr(1, Str, "Barcodelengh", vbTextCompare) Then
          m_Barcodelengh = CInt(Trim(Mid(Str, LenghStr + 1)))
       
      ElseIf InStr(1, Str, "BarcodeChance", vbTextCompare) Then
          m_BarcodeChance = Trim(Mid(Str, LenghStr + 1))
       
      ElseIf InStr(1, Str, "PriceChance", vbTextCompare) Then
          m_PriceChance = Trim(Mid(Str, LenghStr + 1))
       
      ElseIf InStr(1, Str, "RefreshFichNo", vbTextCompare) Then
          m_RefreshFichNo = Trim(Mid(Str, LenghStr + 1))
       
      ElseIf InStr(1, Str, "DirectBascule", vbTextCompare) Then
          m_DirectBascule = Trim(Mid(Str, LenghStr + 1))
          
      ElseIf InStr(1, Str, "StopOnEditFich", vbTextCompare) Then
          m_StopOnEditFich = Trim(Mid(Str, LenghStr + 1))
          
      ElseIf InStr(1, Str, "MultiPrice", vbTextCompare) Then
          m_MultiPrice = Trim(Mid(Str, LenghStr + 1))
          
       ElseIf InStr(1, Str, "UpdateBuyPrice", vbTextCompare) Then
          m_UpdateBuyPrice = Trim(Mid(Str, LenghStr + 1))
          
      ElseIf InStr(1, Str, "UpdateSellPrice", vbTextCompare) Then
          m_UpdateSellPrice = Trim(Mid(Str, LenghStr + 1))
      
      ElseIf InStr(1, Str, "TrazooBarcode", vbTextCompare) Then
          m_TrazooBarcode = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "CodeFlag", vbTextCompare) Then
          m_CodeFlag = Trim(Mid(Str, LenghStr + 1))
          
       ElseIf InStr(1, Str, "CircleMenu", vbTextCompare) Then
          m_CircleMenu = Trim(Mid(Str, LenghStr + 1))
          
       ElseIf InStr(1, Str, "SoundAlarm", vbTextCompare) Then
          m_SoundAlarm = Trim(Mid(Str, LenghStr + 1))
          
       ElseIf InStr(1, Str, "TypeBascule", vbTextCompare) Then
          m_TypeBascule = CInt(Trim(Mid(Str, LenghStr + 1)))
          
       ElseIf InStr(1, Str, "CashPayment", vbTextCompare) Then
          m_CashPayment = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "InpersonSalonPayment", vbTextCompare) Then
          m_InpersonSalonPayment = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "InpersonDeliveryPayment", vbTextCompare) Then
          m_InpersonDeliveryPayment = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "InpersonOutPayment", vbTextCompare) Then
          m_InpersonOutPayment = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "InpersonTablePayment", vbTextCompare) Then
          m_InpersontablePayment = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "InpersonSalonBalance", vbTextCompare) Then
          m_InpersonSalonBalance = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "InpersonDeliveryBalance", vbTextCompare) Then
          m_InpersonDeliveryBalance = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "InpersonOutBalance", vbTextCompare) Then
          m_InpersonOutBalance = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "InpersonTableBalance", vbTextCompare) Then
          m_InpersontableBalance = Trim(Mid(Str, LenghStr + 1))
          
       ElseIf InStr(1, Str, "ByPhoneSalonPayment", vbTextCompare) Then
          m_ByPhoneSalonPayment = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "ByPhoneDeliveryPayment", vbTextCompare) Then
          m_ByPhoneDeliveryPayment = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "ByPhoneTablePayment", vbTextCompare) Then
          m_ByPhonetablePayment = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "ByPhoneSalonBalance", vbTextCompare) Then
          m_ByPhoneSalonBalance = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "ByPhoneDeliveryBalance", vbTextCompare) Then
          m_ByPhoneDeliveryBalance = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "ByPhoneTableBalance", vbTextCompare) Then
          m_ByPhonetableBalance = Trim(Mid(Str, LenghStr + 1))
                 
       ElseIf InStr(1, Str, "ThreeSegmentSearch", vbTextCompare) Then
          m_ThreeSegmentSearch = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "MaxPrices", vbTextCompare) Then
          m_MaxPrices = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "NumberOfUnitSale", vbTextCompare) Then
          m_NumberOfUnitSale = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "MenuViewAfterGood", vbTextCompare) Then
          m_MenuViewAfterGood = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "PayFactorView", vbTextCompare) Then
          m_PayFactorView = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "FromStoreFee", vbTextCompare) Then
          m_FromStoreFee = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "BarcodeAutoEscape", vbTextCompare) Then
          m_BarcodeAutoEscape = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "PrintAfterTasvieh", vbTextCompare) Then
          m_PrintAfterTasvieh = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "AutoBarcode", vbTextCompare) Then
          m_AutoBarcode = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "AutoCallerId", vbTextCompare) Then
          m_AutoCallerId = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "CustomerFarsi", vbTextCompare) Then
          m_CustomerFarsi = Trim(Mid(Str, LenghStr + 1))
       
        ElseIf InStr(1, Str, "CustomerOnlinePrice", vbTextCompare) Then
          m_CustomerOnlinePrice = Trim(Mid(Str, LenghStr + 1))
          
       ElseIf InStr(1, Str, "CustomerAscii", vbTextCompare) Then
          m_CustomerAscii = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "NoCurrentDay", vbTextCompare) Then
          m_NoCurrentDay = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "SaleStartDefault", vbTextCompare) Then
          m_SaleStartDefault = CInt(Trim(Mid(Str, LenghStr + 1)))
    
       ElseIf InStr(1, Str, "ShowDigitNumber", vbTextCompare) Then
          m_ShowDigitNumber = CInt(Trim(Mid(Str, LenghStr + 1)))
       
        ElseIf InStr(1, Str, "CustomerRate", vbTextCompare) Then
          m_CustomerRate = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "GoodPercentage", vbTextCompare) Then
          m_GoodPercentage = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "SearchOrderType", vbTextCompare) Then
          m_SearchOrderType = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "CallerIdSpace", vbTextCompare) Then
          m_CallerIdSpace = Trim(Mid(Str, LenghStr + 1))
       
'''       ElseIf InStr(1, str, "Callwaiting", vbTextCompare) Then
'''          m_Callwaiting = Trim(Mid(str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "CountCustomerDailyBuy", vbTextCompare) Then
          m_CountCustomerDailyBuy = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "CountCustomerGood", vbTextCompare) Then
          m_CountCustomerGood = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "InvoiceRows", vbTextCompare) Then
          m_InvoiceRows = Trim(Mid(Str, LenghStr + 1))
      
       ElseIf InStr(1, Str, "PurchaseRows", vbTextCompare) Then
          m_PurchaseRows = Trim(Mid(Str, LenghStr + 1))
      
       ElseIf InStr(1, Str, "FinalCheck", vbTextCompare) Then
          m_FinalCheck = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "AutoTip", vbTextCompare) Then
          m_AutoTip = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "FichStatusBar", vbTextCompare) Then
          m_FichStatusBar = Trim(Mid(Str, LenghStr + 1))
      
       ElseIf InStr(1, Str, "CycleStockNoDefault", vbTextCompare) Then
          m_CycleStockNoDefault = Val(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "TextIconViewH", vbTextCompare) Then
          m_TextIconViewH = Trim(Mid(Str, LenghStr + 1))
           
       ElseIf InStr(1, Str, "TextIconViewV", vbTextCompare) Then
          m_TextIconViewv = Trim(Mid(Str, LenghStr + 1))
      
       ElseIf InStr(1, Str, "ForceSeller", vbTextCompare) Then
          m_ForceSeller = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "InvoiceStatusDefault", vbTextCompare) Then
          m_InvoiceStatusDefault = Trim(Mid(Str, LenghStr + 1))
      
       ElseIf InStr(1, Str, "NumberOfId", vbTextCompare) Then
          m_NumberOfId = CStr(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "DiscoveryPort", vbTextCompare) Then
          m_DiscoveryPort = CStr(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "ResponsePort", vbTextCompare) Then
          m_ResponsePort = CStr(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "CityCode", vbTextCompare) Then
          m_CityCode = CStr(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "AlmLogFile", vbTextCompare) Then
          m_AlmLogFile = Trim(Mid(Str, LenghStr + 1))
          
       ElseIf InStr(1, Str, "NetworkCallerId", vbTextCompare) Then
          m_NetworkCallerId = Trim(Mid(Str, LenghStr + 1))
      
       ElseIf InStr(1, Str, "TzWeightBarcodeId", vbTextCompare) Then
          m_TzWeightBarcodeId = CStr(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "GoodLevelAutoDiscount", vbTextCompare) Then
          m_GoodLevelAutoDiscount = Trim(Mid(Str, LenghStr + 1))
      
       ElseIf InStr(1, Str, "StartNumberCartReader", vbTextCompare) Then
          m_StartNumberCartReader = CStr(Trim(Mid(Str, LenghStr + 1)))
          
       ElseIf InStr(1, Str, "NumberOfCardReader", vbTextCompare) Then
          m_NumberOfCardReader = CStr(Trim(Mid(Str, LenghStr + 1)))
      
       ElseIf InStr(1, Str, "ReadDirectWeightTz1", vbTextCompare) Then
          m_ReadDirectWeightTz1 = Trim(Mid(Str, LenghStr + 1))
          
       ElseIf InStr(1, Str, "CustomerFeeDataBase", vbTextCompare) Then
          m_CustomerFeeDataBase = Trim(Mid(Str, LenghStr + 1))
                 
       ElseIf InStr(1, Str, "TaxView", vbTextCompare) Then
          m_TaxView = Trim(Mid(Str, LenghStr + 1))
      
       ElseIf InStr(1, Str, "OutPrice", vbTextCompare) Then
          m_OutPrice = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "TzNumericBarcodeId", vbTextCompare) Then
          m_TzNumericBarcodeId = CStr(Trim(Mid(Str, LenghStr + 1)))
       
       ElseIf InStr(1, Str, "NumberOfUnitBuy", vbTextCompare) Then
          m_NumberOfUnitBuy = Trim(Mid(Str, LenghStr + 1))
          
       ElseIf InStr(1, Str, "EditCompatibleSamar1", vbTextCompare) Then
          m_EditCompatibleSamar1 = Trim(Mid(Str, LenghStr + 1))
          
       ElseIf InStr(1, Str, "UndoRedoCompatibleSamar1", vbTextCompare) Then
          m_UndoRedoCompatibleSamar1 = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "EscNotExit", vbTextCompare) Then
          m_EscNotExit = Trim(Mid(Str, LenghStr + 1))
        
       ElseIf InStr(1, Str, "ViewTempAddress", vbTextCompare) Then
          m_ViewTempAddress = Trim(Mid(Str, LenghStr + 1))
          
        ElseIf InStr(1, Str, "CancelNotExit", vbTextCompare) Then
          m_CancelNotExit = Trim(Mid(Str, LenghStr + 1))
       
        ElseIf InStr(1, Str, "ResiveFormerFichCurrentDay", vbTextCompare) Then
          m_ResiveFormerFichCurrentDay = Trim(Mid(Str, LenghStr + 1))
          
        ElseIf InStr(1, Str, "SelectInventory", vbTextCompare) Then
          m_SelectInventory = Trim(Mid(Str, LenghStr + 1))
          
        ElseIf InStr(1, Str, "AlphabetGoodSearch", vbTextCompare) Then
          m_AlphabetGoodSearch = Trim(Mid(Str, LenghStr + 1))
        
        ElseIf InStr(1, Str, "DeficitLog", vbTextCompare) Then
          m_DeficitLog = Trim(Mid(Str, LenghStr + 1))
         
        ElseIf InStr(1, Str, "CallerId8Port", vbTextCompare) Then
          m_CallerId8Port = Trim(Mid(Str, LenghStr + 1))
          
        ElseIf InStr(1, Str, "ShiftRate", vbTextCompare) Then
          m_ShiftRate = Trim(Mid(Str, LenghStr + 1))
          
        ElseIf InStr(1, Str, "SellerCaption", vbTextCompare) Then
          m_SellerCaption = Trim(Mid(Str, LenghStr + 1))
          
        ElseIf InStr(1, Str, "SelectSeller", vbTextCompare) Then
          m_SelectSeller = Trim(Mid(Str, LenghStr + 1))
        
        ElseIf InStr(1, Str, "CountCustomerShiftBuy", vbTextCompare) Then
          m_CountCustomerShiftBuy = Trim(Mid(Str, LenghStr + 1))
          
        ElseIf InStr(1, Str, "FixRateChange", vbTextCompare) Then
          m_FixRateChange = Trim(Mid(Str, LenghStr + 1))
        
        ElseIf InStr(1, Str, "AutoCashClose", vbTextCompare) Then
          m_AutoCashClose = Trim(Mid(Str, LenghStr + 1))
          
       ElseIf InStr(1, Str, "PrintAfterDeliver", vbTextCompare) Then
          m_PrintAfterDeliver = Trim(Mid(Str, LenghStr + 1))
          
       ElseIf InStr(1, Str, "UpDateCarryFee", vbTextCompare) Then
          m_UpDateCarryFee = Trim(Mid(Str, LenghStr + 1))
          
       ElseIf InStr(1, Str, "ShowClock", vbTextCompare) Then
          m_ShowClock = Trim(Mid(Str, LenghStr + 1))
        
        ElseIf InStr(1, Str, "VoiceRecord", vbTextCompare) Then
            m_VoiceRecord = CBool(Trim(Mid(Str, LenghStr + 1)))
        
        ElseIf InStr(1, Str, "TreeViewMenu", vbTextCompare) Then
            m_treeViewMenu = CBool(Trim(Mid(Str, LenghStr + 1)))
        
        ElseIf InStr(1, Str, "CallerIdAutoView", vbTextCompare) Then
            m_CallerIdAutoView = CBool(Trim(Mid(Str, LenghStr + 1)))
        
        ElseIf InStr(1, Str, "CallerIdTest", vbTextCompare) Then
            m_CallerIdTest = CBool(Trim(Mid(Str, LenghStr + 1)))
        
        ElseIf InStr(1, Str, "Pager", vbTextCompare) Then
            m_Pager = CBool(Trim(Mid(Str, LenghStr + 1)))
        
        ElseIf InStr(1, Str, "AutoBackup", vbTextCompare) Then
            m_AutoBackup = Mid(Str, LenghStr + 1)
            
        ElseIf InStr(1, Str, "PosPayment", vbTextCompare) Then
            m_PosPayment = CBool(Mid(Str, LenghStr + 1))
            
        ElseIf InStr(1, Str, "PassPhrase", vbTextCompare) Then
            m_PassPhrase = Mid(Str, LenghStr + 1)
          
        ElseIf InStr(1, Str, "PosApprovedText", vbTextCompare) Then
            m_PosApprovedText = Mid(Str, LenghStr + 1)
        
        ElseIf InStr(1, Str, "PosModel", vbTextCompare) Then
            m_PosModel = Mid(Str, LenghStr + 1)
            
        ElseIf InStr(1, Str, "ReportHeadername", vbTextCompare) Then
            m_ReportHeadername = Mid(Str, LenghStr + 1)
            
        ElseIf InStr(1, Str, "FastCustSave", vbTextCompare) Then
            m_FastCustSave = Mid(Str, LenghStr + 1)
            
       ElseIf InStr(1, Str, "RepetitiveGood", vbTextCompare) Then
          m_RepetitiveGood = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "OtherPartition", vbTextCompare) Then
          m_OtherPartition = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "NotShowPrintNotice", vbTextCompare) Then
          m_NotShowPrintNotice = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "MultiInventory", vbTextCompare) Then
          m_MultiInventory = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Frame_Printers", vbTextCompare) Then
          m_Frame_Printers = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "HasOptionPrice", vbTextCompare) Then
          m_HasOptionPrice = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ShowOption", vbTextCompare) Then
          m_ShowOption = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "LableUsedGood", vbTextCompare) Then
          m_LableUsedGood = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "LoyaltyCustomers", vbTextCompare) Then
          m_LoyaltyCustomers = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "LoyaltyAllCustomers", vbTextCompare) Then
          m_LoyaltyAllCustomers = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "StartCharacter", vbTextCompare) Then
          m_StartCharacter = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "OneClickShow", vbTextCompare) Then
          m_OneClickShow = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "NoRowMenu", vbTextCompare) Then
          m_NoRowMenu = Mid(Str, LenghStr + 1)
       
        ElseIf InStr(1, Str, "RfidReader", vbTextCompare) Then
            m_RfidReader = Mid(Str, LenghStr + 1)
            
        ElseIf InStr(1, Str, "RfidInterval", vbTextCompare) Then
            m_RfidInterval = Mid(Str, LenghStr + 1)
            
        ElseIf InStr(1, Str, "RfidLongBuzzer", vbTextCompare) Then
            m_RfidLongBuzzer = Mid(Str, LenghStr + 1)
            
        ElseIf InStr(1, Str, "TelNetServerActive", vbTextCompare) Then
            m_TelNetServerActive = Mid(Str, LenghStr + 1)
            
        ElseIf InStr(1, Str, "TelNetServerIP", vbTextCompare) Then
            m_TelNetServerIP = Mid(Str, LenghStr + 1)
            
        ElseIf InStr(1, Str, "TelNetServerPort", vbTextCompare) Then
            m_TelNetServerPort = Mid(Str, LenghStr + 1)
            
       ElseIf InStr(1, Str, "PrintAfterPayk", vbTextCompare) Then
          m_PrintAfterPayk = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "PrintAfterOrder", vbTextCompare) Then
          m_PrintAfterOrder = Trim(Mid(Str, LenghStr + 1))
       
        ElseIf InStr(1, Str, "LabelPrint", vbTextCompare) Then
          m_LabelPrint = Mid(Str, LenghStr + 1)
            
       ElseIf InStr(1, Str, "AryaSmsPanel", vbTextCompare) Then
          m_AryaSmsPanel = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ForceTax", vbTextCompare) Then
          m_ForceTax = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "PersonIdCheck", vbTextCompare) Then
          m_PersonIdCheck = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "PersonIdRefreshTime", vbTextCompare) Then
          m_PersonIdRefreshTime = CStr(Trim(Mid(Str, LenghStr + 1)))
      
       ElseIf InStr(1, Str, "ListAutoLoad", vbTextCompare) Then
          m_ListAutoLoad = Mid(Str, LenghStr + 1)
      
       ElseIf InStr(1, Str, "DeviceIP", vbTextCompare) Then
          m_DeviceIP = CStr(Trim(Mid(Str, LenghStr + 1)))
      
       ElseIf InStr(1, Str, "DeviceID", vbTextCompare) Then
          m_DeviceID = CStr(Trim(Mid(Str, LenghStr + 1)))
      
       ElseIf InStr(1, Str, "ListFont", vbTextCompare) Then
          m_ListFont = CStr(Trim(Mid(Str, LenghStr + 1)))
      
       ElseIf InStr(1, Str, "Device2IP", vbTextCompare) Then
          m_Device2IP = CStr(Trim(Mid(Str, LenghStr + 1)))
      
       ElseIf InStr(1, Str, "Device2Id", vbTextCompare) Then
          m_Device2ID = CStr(Trim(Mid(Str, LenghStr + 1)))
      
        End If
    Loop
    tempstring.Close

    If m_CallerIdSpace = 0 Then
        m_CallerIdSpace = 1
    End If
    If m_NoRowMenu = 0 Then m_NoRowMenu = 1
    m_DrawerModel = 0
    m_CustDisplayModel = 0
    m_BasculeModel = 0
    m_CustDisplayPrn = False
    m_CashDrawerPrn = False
    m_BasculeOn = False    ' Default =  Bascule : Off
    m_CashClose = False  ' Cash Open
    m_AccCashClose = False ' Accounting Cash Open
    mvarStartRate = m_PriceType
    If m_ReportHeadername = "" Then m_ReportHeadername = "úÂÑíÇ"
End Sub





