VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mfrmInvoice As Boolean
Private mfrmPurchase As Boolean
Private mfrmGroupAccounting As Boolean
Private mfrmGroupBoxTo As Boolean
Private mfrmPayk As Boolean
Private mfrmPaykPayment As Boolean
Private mfrmGroupPer As Boolean
Private mfrmGroupGoods As Boolean
Private mfrmGroupSystem As Boolean
Private mfrmGroupUtility As Boolean
Private mfrmReports As Boolean
Private mfrmGroupView As Boolean
Private mfrmUserManager As Boolean
Private mfrmManageStore As Boolean
Private mfrmCycleStock As Boolean
Private mfrmCodingGood As Boolean
Private mfrmBackup As Boolean
Private mfrmSms As Boolean
Private mfrmDeviceSetting As Boolean
Private mfrmCust As Boolean
Private mfrmCustComp As Boolean
Private mfrmCustomerDisplay As Boolean
Private mfrmFacEdit As Boolean
Private mfrmFacRecursive As Boolean
Private mfrmFind As Boolean
Private mfrmGood As Boolean
Private mfrmKB As Boolean
Private mfrmKeyBoardTZ1 As Boolean
Private mFrmLogin As Boolean
Private mfrmMenu As Boolean
Private mfrmMenuGoodFirst As Boolean
Private mfrmNotice As Boolean
Private mfrmStationSetting As Boolean
Private mfrmPer As Boolean
Private mfrmPrinter As Boolean
Private mfrmRep As Boolean
Private mfrmRestore As Boolean
Private mfrmSupplier As Boolean
Private mfrmUsePercent As Boolean
Private mfrmSupplierComp As Boolean
Private mfrmWorkTime As Boolean
Private mmdifrm As Boolean
Private mfrmSeller As Boolean
Private mfrmGarson As Boolean
Private mfrmTable As Boolean
Private mfrmGoodPrintFormats As Boolean
Private mfrmGoodDifferences As Boolean
Private mfrmBank As Boolean
Private mfrmSalMali As Boolean
Private mfrmTagheirSalMali As Boolean
Private mfrmFactorReceived As Boolean
Private mfrmCredit As Boolean
Private mfrmCreditCustomer As Boolean
Private mEditInvoice As Boolean
Private mEditPurchase As Boolean
Private mRefferInvoice As Boolean
Private mRefferPurchase As Boolean
Private mFrmAccessLevel As Boolean
Private mfrmPocketPcGroupsAndStations As Boolean
Private mfrmPocketPCGroupGood As Boolean
Private mcarryfee As Boolean
Private mdiscount As Boolean
Private mEditAllStationsFactors As Boolean
Private mpacking As Boolean
Private mRefferedAllStationsFactors As Boolean
Private mservice As Boolean
Private mviewallstationsfactors As Boolean
Private mfrmHistory As Boolean
Private mfrmLablePrinter As Boolean
Private mfrmBranch As Boolean
Private mPrinting As Boolean
Private mReprint As Boolean
Private mInvoicePrint As Boolean
Private mDailyReport As Boolean
Private mCustomizeFee As Boolean
Private mfrmGroupStore As Boolean
Private mfrmCustomReport As Boolean
Private mfrmCreateSanad As Boolean
Private mfrmGoodInKitchen As Boolean
Private mfrmMojodiControl_1 As Boolean
Private mfrmMojodiControl_2 As Boolean
Private mfrmMojodiControl_3 As Boolean
Private mfrmMojodiControl_4 As Boolean
Private mfrmRemainingControl As Boolean
Private mRepGroupGoodsSale As Boolean
Private mRepTimePerSell As Boolean
Private mRepDetailGoodsSale As Boolean
Private mRepBedeh As Boolean
Private mRepBedehDetail As Boolean
Private mRepFichSale As Boolean
Private mRepServeKindSell As Boolean
Private mRepCustPricDiscount As Boolean
Private mRepCustPrice As Boolean
Private mRepSystemGroup As Boolean
Private mRepSerialFich As Boolean
Private mRepShift As Boolean
Private mRepFichBuy As Boolean
Private mRepDetailGoodsBuy As Boolean
Private mRepGroupGoodsBuy As Boolean
Private mRepMojodi As Boolean
Private mRepGoodList As Boolean
Private mRepGarson As Boolean
Private mRepCashInvoice As Boolean
Private mRepStationSaleSummery As Boolean
Private mRepCustomerList As Boolean
Private mRepInventoryRecipt As Boolean
Private mRepGoodDifferences As Boolean
Private mRepUsedGoodAmount As Boolean
Private mRepLossGoodAmount As Boolean
Private mRepStationSaleSummaryByUser As Boolean
Private mRepGetOrderGoodAmount As Boolean
Private mfrmPayment As Boolean
Private mMultiPrice As Boolean
Private mfrmRecieved As Boolean
Private mfrmAccCoding As Boolean
Private mfrmCashClose As Boolean
Private mfrmGoodDiscount As Boolean
Private mfrmCreditSupplier As Boolean
Private mRepDailyWeeding As Boolean
Private mRepDailyPrize As Boolean
Private mReturnFromAccountDeliver As Boolean
Private mRepSaleShiftDailyPrize As Boolean
Private mfrmPhoneBook As Boolean
Private mEditAllFichUser As Boolean
Private mRefferedAllFichUser As Boolean
Private mfrmDistance As Boolean
Private mRepTableSellDetail As Boolean
Private mfrmWeeding As Boolean
Private mfrmGoodLevelAutoDiscount As Boolean
Private mfrmPrize As Boolean
Private mfrmPrizeType As Boolean
Private mfrmInventory As Boolean
Private mfrmInventory_Level1 As Boolean
Private mfrmStation_Inventory As Boolean
Private mfrmDoreAuction As Boolean
Private mfrmLevel2Discount As Boolean
Private mfrmDoreAuctionRuleDiscount As Boolean
Private mfrmColorSizeBrandSeason As Boolean
Private mfrmLottery As Boolean
Private mfrmDoreLottery As Boolean
Private mfrmAutoDiscount As Boolean
Private mfrmGoodPer As Boolean
Private mRepSeller As Boolean
Private mfrmManageSet As Boolean
Private mfrmCustGoodTurnOver As Boolean
Private mfrmGoodPicture As Boolean
Private mEditInvoiceCashClose As Boolean
Private mMojodiControl As Boolean
Private mfrmAccount As Boolean
Private mUpperAmountGood As Boolean
Private mSaveWithoutPrint As Boolean
Private mCashRemaining As Boolean
Private mNavigateFactor As Boolean
Private mfrmFindFactor As Boolean
Private mfrmTaxAndDuty As Boolean
Private mfrmTestPos As Boolean
Private mLockCheck As Boolean
Private mReduceAmountGood As Boolean
Private mfrmPos As Boolean
Private mUnBalance As Boolean
Private mGraphicTables As Boolean
Private mfrmGoodTurnOver As Boolean
Private mfrmCreditCustomerAccount As Boolean
Private mChangeTotalRemainingAmount As Boolean
Private mAccfrmAsnad As Boolean
Private mAccfrmCheckBook As Boolean
Private mSalarySystem As Boolean
Private mAccfrmKartHesabReport As Boolean
Private mfrmTarazSoodZian As Boolean
Private mFullSummaryReportSaved As Boolean
Private mfrmBenefit As Boolean
Private mfrmTempFich As Boolean
Private mfrmHavaleh As Boolean
Private mfrmTempReceived As Boolean
Private mfrmSavePurchase As Boolean
Private mfrmLosses As Boolean
Private mfrmSaveTempReceived As Boolean
Private mFirstMojodiControl As Boolean
Private mfrmOrder As Boolean
Private mPriceFactorDecrease As Boolean
Private mSaleReturn As Boolean

Dim Parameter() As Parameter

Public Property Get EditInvoice() As Boolean
    EditInvoice = mEditInvoice
End Property

Public Property Get EditPurchase() As Boolean
    EditPurchase = mEditPurchase
End Property

Public Property Get RefferInvoice() As Boolean
    RefferInvoice = mRefferInvoice
End Property

Public Property Get RefferPurchase() As Boolean
    RefferPurchase = mRefferPurchase
End Property

Public Property Get frmInvoice() As Boolean
    frmInvoice = mfrmInvoice
End Property

Public Property Get frmPurchase() As Boolean
    frmPurchase = mfrmPurchase
End Property

Public Property Get frmGroupAccounting() As Boolean
    frmGroupAccounting = mfrmGroupAccounting
End Property

Public Property Get frmGroupBoxTo() As Boolean
    frmGroupBoxTo = mfrmGroupBoxTo
End Property

Public Property Get frmPayk() As Boolean
    frmPayk = mfrmPayk
End Property

Public Property Get frmPaykPayment() As Boolean
    frmPaykPayment = mfrmPaykPayment
End Property

Public Property Get frmGroupPer() As Boolean
    frmGroupPer = mfrmGroupPer
End Property

Public Property Get frmGroupGoods() As Boolean
    frmGroupGoods = mfrmGroupGoods
End Property

Public Property Get frmGroupSystem() As Boolean
    frmGroupSystem = mfrmGroupSystem
End Property

Public Property Get frmGroupUtility() As Boolean
    frmGroupUtility = mfrmGroupUtility
End Property

Public Property Get frmReports() As Boolean
    frmReports = mfrmReports
End Property

Public Property Get frmGroupView() As Boolean
    frmGroupView = mfrmGroupView
End Property

Public Property Get frmUserManager() As Boolean
    frmUserManager = mfrmUserManager
End Property

Public Property Get frmCodingGood() As Boolean
    frmCodingGood = mfrmCodingGood
End Property

Public Property Get frmBackup() As Boolean
    frmBackup = mfrmBackup
End Property
Public Property Get frmSms() As Boolean
    frmSms = mfrmSms
End Property
Public Property Get frmDeviceSetting() As Boolean
    frmDeviceSetting = mfrmDeviceSetting
End Property

Public Property Get frmCust() As Boolean
    frmCust = mfrmCust
End Property

Public Property Get frmCustComp() As Boolean
    frmCustComp = mfrmCustComp
End Property

Public Property Get frmCustomerDisplay() As Boolean
    frmCustomerDisplay = mfrmCustomerDisplay
End Property

'Public Property Get FrmDrawerModel() As Boolean
'    FrmDrawerModel = mFrmDrawerModel
'End Property

Public Property Get frmFacEdit() As Boolean
    frmFacEdit = mfrmFacEdit
End Property

Public Property Get frmFacRecursive() As Boolean
    frmFacRecursive = mfrmFacRecursive
End Property

Public Property Get frmFind() As Boolean
    frmFind = mfrmFind
End Property

Public Property Get frmGood() As Boolean
    frmGood = mfrmGood
End Property

Public Property Get frmKB() As Boolean
    frmKB = mfrmKB
End Property

Public Property Get frmKeyBoardTZ1() As Boolean
    frmKeyBoardTZ1 = mfrmKeyBoardTZ1
End Property

Public Property Get FrmLogin() As Boolean
    FrmLogin = mFrmLogin
End Property

Public Property Get frmMenu() As Boolean
    frmMenu = mfrmMenu
End Property

Public Property Get frmMenuGoodFirst() As Boolean
    frmMenuGoodFirst = mfrmMenuGoodFirst
End Property

Public Property Get frmNotice() As Boolean
    frmNotice = mfrmNotice
End Property

Public Property Get frmStationsetting() As Boolean
    frmStationsetting = mfrmStationSetting
End Property

Public Property Get frmPer() As Boolean
    frmPer = mfrmPer
End Property

Public Property Get frmPrinter() As Boolean
    frmPrinter = mfrmPrinter
End Property

Public Property Get frmRep() As Boolean
    frmRep = mfrmRep
End Property

Public Property Get frmRestore() As Boolean
    frmRestore = mfrmRestore
End Property

Public Property Get frmSupplier() As Boolean
    frmSupplier = mfrmSupplier
End Property

Public Property Get frmUsePercent() As Boolean
    frmUsePercent = mfrmUsePercent
End Property


Public Property Get frmSupplierComp() As Boolean
    frmSupplierComp = mfrmSupplierComp
End Property

Public Property Get frmWorkTime() As Boolean
    frmWorkTime = mfrmWorkTime
End Property

Public Property Get frmGoodDifferences() As Boolean
    frmGoodDifferences = mfrmGoodDifferences
End Property

Public Property Get mdifrm() As Boolean
    mdifrm = mmdifrm
End Property

Public Property Get frmSeller() As Boolean
    frmSeller = mfrmSeller
End Property
Public Property Get frmGarson() As Boolean
    frmGarson = mfrmGarson
End Property

Public Property Get frmTable() As Boolean
    frmTable = mfrmTable
End Property
Public Property Get frmGoodPrintFormats() As Boolean
    frmGoodPrintFormats = mfrmGoodPrintFormats
End Property

Public Property Get frmBank() As Boolean
    frmBank = mfrmBank
End Property

Public Property Get frmSalMali() As Boolean
    frmSalMali = mfrmSalMali
End Property

Public Property Get frmTagheirSalMali() As Boolean
    frmTagheirSalMali = mfrmTagheirSalMali
End Property

Public Property Get frmFactorReceived() As Boolean
    frmFactorReceived = mfrmFactorReceived
End Property

Public Property Get frmCredit() As Boolean
    frmCredit = mfrmCredit
End Property

Public Property Get frmCreditCustomer() As Boolean
    frmCreditCustomer = mfrmCreditCustomer
End Property

Public Property Get FrmAccessLevel() As Boolean
    FrmAccessLevel = mFrmAccessLevel
End Property

Public Property Get frmPocketPcGroupsAndStations() As Boolean
    frmPocketPcGroupsAndStations = mfrmPocketPcGroupsAndStations
End Property

Public Property Get frmPocketPCGroupGood() As Boolean
    frmPocketPCGroupGood = mfrmPocketPCGroupGood
End Property

Public Property Get carryfee() As Boolean
    carryfee = mcarryfee
End Property

Public Property Get Discount() As Boolean
    Discount = mdiscount
End Property

Public Property Get EditAllStationsFactors() As Boolean
    EditAllStationsFactors = mEditAllStationsFactors
End Property

Public Property Get packing() As Boolean
    packing = mpacking
End Property

Public Property Get RefferedAllStationsFactors() As Boolean
    RefferedAllStationsFactors = mRefferedAllStationsFactors
End Property

Public Property Get service() As Boolean
    service = mservice
End Property

Public Property Get viewallstationsfactors() As Boolean
    viewallstationsfactors = mviewallstationsfactors
End Property

Public Property Get frmHistory() As Boolean
    frmHistory = mfrmHistory
End Property

Public Property Get frmLablePrinter() As Boolean
    frmLablePrinter = mfrmLablePrinter
End Property
Public Property Get frmGroupStore() As Boolean
    frmGroupStore = mfrmGroupStore
End Property

Public Property Get frmBranch() As Boolean
    frmBranch = mfrmBranch
End Property

Public Property Get Reprint() As Boolean
    Reprint = mReprint
End Property
Public Property Get InvoicePrint() As Boolean
    InvoicePrint = mInvoicePrint
End Property
Public Property Get Printing() As Boolean
    Printing = mPrinting
End Property
Public Property Get DailyReport() As Boolean
    DailyReport = mDailyReport
End Property
Public Property Get CustomizeFee() As Boolean
    CustomizeFee = mCustomizeFee
End Property
Public Property Get frmCustomReport() As Boolean
    frmCustomReport = mfrmCustomReport
End Property
Public Property Get frmCreateSanad() As Boolean
    frmCreateSanad = mfrmCreateSanad
End Property

Public Property Get frmGoodInKitchen() As Boolean
    frmGoodInKitchen = mfrmGoodInKitchen
End Property
Public Property Get frmMojodiControl_1() As Boolean
    frmMojodiControl_1 = mfrmMojodiControl_1
End Property
Public Property Get frmMojodiControl_2() As Boolean
    frmMojodiControl_2 = mfrmMojodiControl_2
End Property
Public Property Get frmMojodiControl_3() As Boolean
    frmMojodiControl_3 = mfrmMojodiControl_3
End Property
Public Property Get frmMojodiControl_4() As Boolean
    frmMojodiControl_4 = mfrmMojodiControl_4
End Property
Public Property Get frmRemainingControl() As Boolean
    frmRemainingControl = mfrmRemainingControl
End Property
Public Property Get frmManageStore() As Boolean
    frmManageStore = mfrmManageStore
End Property
Public Property Get frmCycleStock() As Boolean
   frmCycleStock = mfrmCycleStock
End Property
Public Property Get RepGroupGoodsSale() As Boolean
    RepGroupGoodsSale = mRepGroupGoodsSale
End Property
Public Property Get RepTimePerSell() As Boolean
    RepTimePerSell = mRepTimePerSell
End Property
Public Property Get RepDetailGoodsSale() As Boolean
    RepDetailGoodsSale = mRepDetailGoodsSale
End Property
Public Property Get RepBedeh() As Boolean
    RepBedeh = mRepBedeh
End Property
Public Property Get RepBedehDetail() As Boolean
    RepBedehDetail = mRepBedehDetail
End Property
Public Property Get RepFichSale() As Boolean
    RepFichSale = mRepFichSale
End Property
Public Property Get RepServeKindSell() As Boolean
    RepServeKindSell = mRepServeKindSell
End Property
Public Property Get RepCustPricDiscount() As Boolean
    RepCustPricDiscount = mRepCustPricDiscount
End Property
Public Property Get RepCustPrice() As Boolean
    RepCustPrice = mRepCustPrice
End Property
Public Property Get RepSystemGroup() As Boolean
    RepSystemGroup = mRepSystemGroup
End Property
Public Property Get RepSerialFich() As Boolean
    RepSerialFich = mRepSerialFich
End Property
Public Property Get RepShift() As Boolean
    RepShift = mRepShift
End Property
Public Property Get RepFichBuy() As Boolean
    RepFichBuy = mRepFichBuy
End Property
Public Property Get RepDetailGoodsBuy() As Boolean
    RepDetailGoodsBuy = mRepDetailGoodsBuy
End Property
Public Property Get RepGroupGoodsBuy() As Boolean
    RepGroupGoodsBuy = mRepGroupGoodsBuy
End Property
Public Property Get RepMojodi() As Boolean
    RepMojodi = mRepMojodi
End Property
Public Property Get RepGoodList() As Boolean
    RepGoodList = mRepGoodList
End Property
Public Property Get RepGarson() As Boolean
    RepGarson = mRepGarson
End Property
Public Property Get RepCashInvoice() As Boolean
    RepCashInvoice = mRepCashInvoice
End Property
Public Property Get RepStationSaleSummery() As Boolean
    RepStationSaleSummery = mRepStationSaleSummery
End Property
Public Property Get RepCustomerList() As Boolean
    RepCustomerList = mRepCustomerList
End Property
Public Property Get RepInventoryRecipt() As Boolean
    RepInventoryRecipt = mRepInventoryRecipt
End Property
Public Property Get RepGoodDifferences() As Boolean
    RepGoodDifferences = mRepGoodDifferences
End Property
Public Property Get RepUsedGoodAmount() As Boolean
    RepUsedGoodAmount = mRepUsedGoodAmount
End Property
Public Property Get RepLossGoodAmount() As Boolean
    RepLossGoodAmount = mRepLossGoodAmount
End Property
Public Property Get RepStationSaleSummaryByUser() As Boolean
    RepStationSaleSummaryByUser = mRepStationSaleSummaryByUser
End Property
Public Property Get RepGetOrderGoodAmount() As Boolean
    RepGetOrderGoodAmount = mRepGetOrderGoodAmount
End Property
Public Property Get frmPayment() As Boolean
    frmPayment = mfrmPayment
End Property
Public Property Get MultiPrice() As Boolean
    MultiPrice = mMultiPrice
End Property
Public Property Get frmRecieved() As Boolean
    frmRecieved = mfrmRecieved
End Property
Public Property Get frmAccCoding() As Boolean
    frmAccCoding = mfrmAccCoding
End Property
Public Property Get frmCashClose() As Boolean
    frmCashClose = mfrmCashClose
End Property
Public Property Get frmGoodDiscount() As Boolean
    frmGoodDiscount = mfrmGoodDiscount
End Property
Public Property Get frmCreditSupplier() As Boolean
    frmCreditSupplier = mfrmCreditSupplier
End Property
Public Property Get RepDailyWeeding() As Boolean
    RepDailyWeeding = mRepDailyWeeding
End Property
Public Property Get RepDailyPrize() As Boolean
    RepDailyPrize = mRepDailyPrize
End Property
Public Property Get ReturnFromAccountDeliver() As Boolean
    ReturnFromAccountDeliver = mReturnFromAccountDeliver
End Property
Public Property Get RepSaleShiftDailyPrize() As Boolean
    RepSaleShiftDailyPrize = mRepSaleShiftDailyPrize
End Property
Public Property Get frmPhoneBook() As Boolean
    frmPhoneBook = mfrmPhoneBook
End Property
Public Property Get EditAllFichUser() As Boolean
    EditAllFichUser = mEditAllFichUser
End Property
Public Property Get RefferedAllFichUser() As Boolean
    RefferedAllFichUser = mRefferedAllFichUser
End Property
Public Property Get frmDistance() As Boolean
    frmDistance = mfrmDistance
End Property
Public Property Get RepTableSellDetail() As Boolean
    RepTableSellDetail = mRepTableSellDetail
End Property
Public Property Get frmWeeding() As Boolean
    frmWeeding = mfrmWeeding
End Property
Public Property Get frmGoodLevelAutoDiscount() As Boolean
    frmGoodLevelAutoDiscount = mfrmGoodLevelAutoDiscount
End Property
Public Property Get frmPrize() As Boolean
    frmPrize = mfrmPrize
End Property
Public Property Get frmPrizeType() As Boolean
    frmPrizeType = mfrmPrizeType
End Property
Public Property Get frmInventory() As Boolean
    frmInventory = mfrmInventory
End Property
Public Property Get frmInventory_level1() As Boolean
    frmInventory_level1 = mfrmInventory_Level1
End Property
Public Property Get frmStation_Inventory() As Boolean
    frmStation_Inventory = mfrmStation_Inventory
End Property
Public Property Get frmDoreAuction() As Boolean
    frmDoreAuction = mfrmDoreAuction
End Property
Public Property Get frmLevel2Discount() As Boolean
    frmLevel2Discount = mfrmLevel2Discount
End Property
Public Property Get frmDoreAuctionRuleDiscount() As Boolean
    frmDoreAuctionRuleDiscount = mfrmDoreAuctionRuleDiscount
End Property
Public Property Get frmColorSizeBrandSeason() As Boolean
    frmColorSizeBrandSeason = mfrmColorSizeBrandSeason
End Property
Public Property Get frmLottery() As Boolean
    frmLottery = mfrmLottery
End Property
Public Property Get frmDoreLottery() As Boolean
    frmDoreLottery = mfrmDoreLottery
End Property
Public Property Get frmAutodiscount() As Boolean
    frmAutodiscount = mfrmAutoDiscount
End Property
Public Property Get frmGoodPer() As Boolean
    frmGoodPer = mfrmGoodPer
End Property
Public Property Get RepSeller() As Boolean
    RepSeller = mRepSeller
End Property
Public Property Get frmManageSet() As Boolean
    frmManageSet = mfrmManageSet
End Property
Public Property Get frmCustGoodTurnOver() As Boolean
    frmCustGoodTurnOver = mfrmCustGoodTurnOver
End Property
Public Property Get frmGoodPicture() As Boolean
    frmGoodPicture = mfrmGoodPicture
End Property
Public Property Get EditInvoiceCashClose() As Boolean
    EditInvoiceCashClose = mEditInvoiceCashClose
End Property
Public Property Get MojodiControl() As Boolean
    MojodiControl = mMojodiControl
End Property
Public Property Get frmAccount() As Boolean
    frmAccount = mfrmAccount
End Property
Public Property Get UpperAmountGood() As Boolean
    UpperAmountGood = mUpperAmountGood
End Property
Public Property Get SaveWithoutPrint() As Boolean
    SaveWithoutPrint = mSaveWithoutPrint
End Property
Public Property Get CashRemaining() As Boolean
    CashRemaining = mCashRemaining
End Property
Public Property Get NavigateFactor() As Boolean
    NavigateFactor = mNavigateFactor
End Property
Public Property Get frmFindFactor() As Boolean
    frmFindFactor = mfrmFindFactor
End Property
Public Property Get frmTaxAndDuty() As Boolean
    frmTaxAndDuty = mfrmTaxAndDuty
End Property
Public Property Get frmTestPos() As Boolean
    frmTestPos = mfrmTestPos
End Property
Public Property Get LockCheck() As Boolean
    LockCheck = mLockCheck
End Property
Public Property Get ReduceAmountGood() As Boolean
    ReduceAmountGood = mReduceAmountGood
End Property
Public Property Get frmPos() As Boolean
    frmPos = mfrmPos
End Property
Public Property Get UnBalance() As Boolean
    UnBalance = mUnBalance
End Property
Public Property Get GraphicTables() As Boolean
    GraphicTables = mGraphicTables
End Property
Public Property Get frmGoodTurnOver() As Boolean
    frmGoodTurnOver = mfrmGoodTurnOver
End Property
Public Property Get frmCreditCustomerAccount() As Boolean
    frmCreditCustomerAccount = mfrmCreditCustomerAccount
End Property
Public Property Get ChangeTotalRemainingAmount() As Boolean
    ChangeTotalRemainingAmount = mChangeTotalRemainingAmount
End Property
Public Property Get AccfrmAsnad() As Boolean
    AccfrmAsnad = mAccfrmAsnad
End Property
Public Property Get AccfrmCheckBook() As Boolean
    AccfrmCheckBook = mAccfrmCheckBook
End Property
Public Property Get SalarySystem() As Boolean
    SalarySystem = mSalarySystem
End Property
Public Property Get AccfrmKartHesabReport() As Boolean
    AccfrmKartHesabReport = mAccfrmKartHesabReport
End Property
Public Property Get frmTarazSoodZian() As Boolean
    frmTarazSoodZian = mfrmTarazSoodZian
End Property
Public Property Get FullSummaryReportSaved() As Boolean
    FullSummaryReportSaved = mFullSummaryReportSaved
End Property
Public Property Get frmBenefit() As Boolean
    frmBenefit = mfrmBenefit
End Property
Public Property Get frmTempFich() As Boolean
    frmTempFich = mfrmTempFich
End Property
Public Property Get frmTempReceived() As Boolean
    frmTempReceived = mfrmTempReceived
End Property
Public Property Get frmHavaleh() As Boolean
    frmHavaleh = mfrmHavaleh
End Property
Public Property Get frmSavePurchase() As Boolean
    frmSavePurchase = mfrmSavePurchase
End Property
Public Property Get frmLosses() As Boolean
    frmLosses = mfrmLosses
End Property
Public Property Get frmSaveTempReceived() As Boolean
    frmSaveTempReceived = mfrmSaveTempReceived
End Property
Public Property Get FirstMojodiControl() As Boolean
    FirstMojodiControl = mFirstMojodiControl
End Property
Public Property Get frmOrder() As Boolean
    frmOrder = mfrmOrder
End Property
Public Property Get PriceFactorDecrease() As Boolean
    PriceFactorDecrease = mPriceFactorDecrease
End Property
Public Property Get SaleReturn() As Boolean
    SaleReturn = mSaleReturn
End Property

Public Sub Class_Initialize()
    If mVarAccessLevel = 0 Then Exit Sub
    
    Dim ClsCnn As New ADODB.Connection
    Dim Rst As New ADODB.Recordset

    ClsCnn.ConnectionString = strConnectionString
    ClsCnn.Open
    
    ClearAccess
    
    ReDim Parameter(1) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@AccessLevel", adInteger, 4, mVarAccessLevel)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Object_Access", Parameter, ClsCnn)
    
    While Not Rst.EOF
        If IsNull(Rst.Fields("intAccessLevel").Value) = False Then
            Select Case LCase(Rst.Fields("ObjectId").Value)
            
                Case LCase("frmInvoice")
                    mfrmInvoice = True
                    
                Case LCase("frmPurchase")
                    mfrmPurchase = True
                    
                Case LCase("frmGroupAccounting")
                    mfrmGroupAccounting = True
                    
                Case LCase("frmGroupBoxTo")
                    mfrmGroupBoxTo = True
                    
                Case LCase("frmPayk")
                    mfrmPayk = True
                    
                Case LCase("frmPaykPayment")
                    mfrmPaykPayment = True
                    
                Case LCase("frmGroupPer")
                    mfrmGroupPer = True
                    
                Case LCase("frmGroupGoods")
                    mfrmGroupGoods = True
                    
                Case LCase("frmGroupSystem")
                    mfrmGroupSystem = True
                    
                Case LCase("frmGroupUtility")
                    mfrmGroupUtility = True
                    
                Case LCase("frmReports")
                    mfrmReports = True
                    
                Case LCase("frmGroupView")
                    mfrmGroupView = True
                    
                Case LCase("frmCodingGood")
                    mfrmCodingGood = True
                
                Case LCase("frmBackup")
                    mfrmBackup = True
                
                Case LCase("frmSms")
                    mfrmSms = True
                    
                Case LCase("frmDeviceSetting")
                    mfrmDeviceSetting = True
                    
                Case LCase("frmCust")
                    mfrmCust = True
                    
                Case LCase("frmCustComp")
                    mfrmCustComp = True
                    
                Case LCase("frmCustomerDisplay")
                    mfrmCustomerDisplay = True
                                    
                Case LCase("FrmDrawerModel")
                    mFrmDrawerModel = True
                    
                Case LCase("frmFacEdit")
                    mfrmFacEdit = True
                    
                Case LCase("frmFacRecursive")
                    mfrmFacRecursive = True
                    
                Case LCase("frmFind")
                    mfrmFind = True
                    
                Case LCase("frmGood")
                    mfrmGood = True
                    
                Case LCase("frmKB")
                    mfrmKB = True
                    
                Case LCase("frmKeyBoardTZ1")
                    mfrmKeyBoardTZ1 = True
                    
                Case LCase("FrmLogin")
                    mFrmLogin = True
                    
                Case LCase("frmMenu")
                    mfrmMenu = True
                    
                Case LCase("frmMenuGoodFirst")
                    mfrmMenuGoodFirst = True
                    
                Case LCase("frmNotice")
                    mfrmNotice = True
                    
                Case LCase("frmStationSetting")
                    mfrmStationSetting = True
                    
                Case LCase("frmPer")
                    mfrmPer = True
                    
                Case LCase("frmPrinter")
                    mfrmPrinter = True
                    
                Case LCase("frmRep")
                    mfrmRep = True
                    
                Case LCase("frmRestore")
                    mfrmRestore = True
                    
                Case LCase("frmSupplier")
                    mfrmSupplier = True
                    
                Case LCase("frmUsePercent")
                    mfrmUsePercent = True
                    
                Case LCase("frmSupplierComp")
                    mfrmSupplierComp = True
                    
                Case LCase("frmWorkTime")
                    mfrmWorkTime = True
                    
                Case LCase("mdifrm")
                    mmdifrm = True
                    
                Case LCase("frmseller")
                    mfrmSeller = True
                    
                Case LCase("frmGarson")
                    mfrmGarson = True
                    
                Case LCase("frmTable")
                    mfrmTable = True
                
                Case LCase("frmGoodPrintFormats")
                    mfrmGoodPrintFormats = True
                    
                Case LCase("frmGoodDifferences")
                    mfrmGoodDifferences = True
                    
                Case LCase("EditInvoice")
                    mEditInvoice = True
                    
                Case LCase("EditPurchase")
                    mEditPurchase = True
                    
                Case LCase("RefferInvoice")
                    mRefferInvoice = True
                    
                Case LCase("RefferPurchase")
                    mRefferPurchase = True
                
                Case LCase("frmBank")
                    mfrmBank = True
                
                Case LCase("frmSalMali")
                    mfrmSalMali = True
                
                Case LCase("frmTagheirSalMali")
                    mfrmTagheirSalMali = True
                
                Case LCase("frmFactorReceived")
                    mfrmFactorReceived = True
                
                Case LCase("frmCredit")
                    mfrmCredit = True
                    
                Case LCase("frmCreditCustomer")
                    mfrmCreditCustomer = True
                    
                Case LCase("FrmAccessLevel")
                    mFrmAccessLevel = True
                    
                Case LCase("frmPocketPcGroupsAndStations")
                    mfrmPocketPcGroupsAndStations = True
                    
                Case LCase("frmPocketPCGroupGood")
                    mfrmPocketPCGroupGood = True
                    
                Case LCase("carryfee")
                    mcarryfee = True
                    
                Case LCase("discount")
                    mdiscount = True
                    
                Case LCase("EditAllStationsFactors")
                    mEditAllStationsFactors = True
                    
                Case LCase("packing")
                    mpacking = True
                    
                Case LCase("RefferedAllStationsFactors")
                    mRefferedAllStationsFactors = True
                    
                Case LCase("service")
                    mservice = True
                    
                Case LCase("viewallstationsfactors")
                    mviewallstationsfactors = True
                
                Case LCase("frmhistory")
                    mfrmHistory = True
                
                Case LCase("frmlablePrinter")
                    mfrmLablePrinter = True
                    
                Case LCase("frmGroupStore")
                    mfrmGroupStore = True
                    
                Case LCase("frmBranch")
                    mfrmBranch = True
                
                Case LCase("frmsuppliersomp")
                    mfrmSupplierComp = True
                    
                Case LCase("frmsupplier")
                    mfrmSupplier = True
                    
                Case LCase("Printing")
                    mPrinting = True
                    
                Case LCase("Reprint")
                    mReprint = True
                    
                Case LCase("InvoicePrint")
                    mInvoicePrint = True
                    
                Case LCase("DailyReport")
                    mDailyReport = True
                    
                Case LCase("CustomizeFee")
                    mCustomizeFee = True
                    
                Case LCase("frmCustomReport")
                    mfrmCustomReport = True
                    
                Case LCase("frmCreateSanad")
                    mfrmCreateSanad = True
                    
                Case LCase("frmGoodInKitchen")
                    mfrmGoodInKitchen = True
                
                Case LCase("frmMojodiControl_1")
                    mfrmMojodiControl_1 = True
                    
                Case LCase("frmMojodiControl_2")
                    mfrmMojodiControl_2 = True
                    
                Case LCase("frmMojodiControl_3")
                    mfrmMojodiControl_3 = True
                    
                Case LCase("frmMojodiControl_4")
                    mfrmMojodiControl_4 = True
                    
                Case LCase("frmRemainingControl")
                    mfrmRemainingControl = True
                    
                Case LCase("RepGroupGoodsSale")
                    mRepGroupGoodsSale = True
                    
                Case LCase("RepTimePerSell")
                    mRepTimePerSell = True
                    
                Case LCase("RepDetailGoodsSale")
                    mRepDetailGoodsSale = True
                    
                Case LCase("RepBedeh")
                    mRepBedeh = True
                    
                Case LCase("RepBedehDetail")
                    mRepBedehDetail = True
                    
                Case LCase("RepFichSale")
                    mRepFichSale = True
                    
                Case LCase("RepServeKindSell")
                    mRepServeKindSell = True
                    
                Case LCase("RepCustPricDiscount")
                    mRepCustPricDiscount = True
                    
                Case LCase("RepCustPrice")
                    mRepCustPrice = True
                    
                Case LCase("RepSystemGroup")
                    mRepSystemGroup = True
                    
                Case LCase("RepSerialFich")
                    mRepSerialFich = True
                    
                Case LCase("RepShift")
                    mRepShift = True
                    
                Case LCase("RepFichBuy")
                    mRepFichBuy = True
                    
                Case LCase("RepDetailGoodsBuy")
                    mRepDetailGoodsBuy = True
                    
                Case LCase("RepGroupGoodsBuy")
                    mRepGroupGoodsBuy = True
                    
                Case LCase("RepMojodi")
                    mRepMojodi = True
                    
                Case LCase("RepGoodList")
                    mRepGoodList = True
                    
                Case LCase("RepGarson")
                    mRepGarson = True
                    
                Case LCase("RepCashInvoice")
                    mRepCashInvoice = True
                    
                Case LCase("RepStationSaleSummery")
                    mRepStationSaleSummery = True
                    
                Case LCase("RepCustomerList")
                    mRepCustomerList = True
                    
                Case LCase("RepInventoryRecipt")
                    mRepInventoryRecipt = True
                    
                Case LCase("RepGoodDifferences")
                    mRepGoodDifferences = True
                    
                Case LCase("RepUsedGoodAmount")
                    mRepUsedGoodAmount = True
                    
                Case LCase("RepLossGoodAmount")
                    mRepLossGoodAmount = True
                    
                Case LCase("RepStationSaleSummaryByUser")
                    mRepStationSaleSummaryByUser = True
                    
                Case LCase("RepGetOrderGoodAmount")
                    mRepGetOrderGoodAmount = True
                
                Case LCase("frmPayment")
                    mfrmPayment = True
                    
                Case LCase("MultiPrice")
                    mMultiPrice = True
                    
                Case LCase("frmRecieved")
                    mfrmRecieved = True
                    
                Case LCase("frmAccCoding")
                    mfrmAccCoding = True
                    
                Case LCase("frmCashClose")
                    mfrmCashClose = True
                    
                Case LCase("frmGoodDiscount")
                    mfrmGoodDiscount = True
                    
                Case LCase("frmCreditSupplier")
                    mfrmCreditSupplier = True
                    
                Case LCase("RepDailyWeeding")
                    mRepDailyWeeding = True
                    
                Case LCase("RepDailyPrize")
                    mRepDailyPrize = True
                    
                Case LCase("ReturnFromAccountDeliver")
                    mReturnFromAccountDeliver = True
                    
                Case LCase("RepSaleShiftDailyPrize")
                    mRepSaleShiftDailyPrize = True
                    
                Case LCase("frmPhoneBook")
                    mfrmPhoneBook = True
                    
                Case LCase("EditAllFichUser")
                    mEditAllFichUser = True
                    
                Case LCase("RefferedAllFichUser")
                    mRefferedAllFichUser = True
                
                Case LCase("frmDistance")
                    mfrmDistance = True
                
                Case LCase("RepTableSellDetail")
                    mRepTableSellDetail = True
                    
                Case LCase("frmManageStore")
                    mfrmManageStore = True
                
                Case LCase("frmCycleStock")
                    mfrmCycleStock = True
                    
                Case LCase("frmWeeding")
                    mfrmWeeding = True
                     
                Case LCase("frmGoodLevelAutoDiscount")
                    mfrmGoodLevelAutoDiscount = True
                
                Case LCase("frmPrize")
                    mfrmPrize = True
                    
                Case LCase("frmPrizeType")
                    mfrmPrizeType = True
                    
                Case LCase("frmInventory")
                    mfrmInventory = True
                    
                Case LCase("frmInventory_Level1")
                    mfrmInventory_Level1 = True
                    
                Case LCase("frmStation_Inventory")
                    mfrmStation_Inventory = True
                
                Case LCase("frmDoreAuction")
                    mfrmDoreAuction = True
                 
                 Case LCase("frmLevel2Discount")
                    mfrmLevel2Discount = True
                
                Case LCase("frmDoreAuctionRuleDiscount")
                    mfrmDoreAuctionRuleDiscount = True
                
                Case LCase("frmColorSizeBrandSeason")
                    mfrmColorSizeBrandSeason = True
                
                Case LCase("frmLottery")
                    mfrmLottery = True
            
                Case LCase("frmDoreLottery")
                    mfrmDoreLottery = True
            
                Case LCase("frmAutoDiscount")
                    mfrmAutoDiscount = True
                    
                Case LCase("frmGoodPer")
                    mfrmGoodPer = True
                
                Case LCase("RepSeller")
                    mRepSeller = True
                    
                Case LCase("frmManageSet")
                    mfrmManageSet = True
                
                Case LCase("frmCustGoodTurnOver")
                    mfrmCustGoodTurnOver = True
                    
                Case LCase("frmGoodPicture")
                    mfrmGoodPicture = True
                
                Case LCase("EditInvoiceCashClose")
                    mEditInvoiceCashClose = True
                
                Case LCase("MojodiControl")
                    mMojodiControl = True
                    
                Case LCase("frmAccount")
                    mfrmAccount = True
                
                Case LCase("UpperAmountGood")
                    mUpperAmountGood = True
                
                Case LCase("SaveWithoutPrint")
                    mSaveWithoutPrint = True

                Case LCase("CashRemaining")
                    mCashRemaining = True

                Case LCase("NavigateFactor")
                    mNavigateFactor = True

                Case LCase("frmFindFactor")
                    mfrmFindFactor = True

                Case LCase("frmTaxAndDuty")
                    mfrmTaxAndDuty = True

                Case LCase("frmTestPos")
                    mfrmTestPos = True

                Case LCase("LockCheck")
                    mLockCheck = True
                    
                Case LCase("ReduceAmountGood")
                    mReduceAmountGood = True
                    
                Case LCase("frmPos")
                    mfrmPos = True
                    
                Case LCase("UnBalance")
                    mUnBalance = True
                    
                Case LCase("GraphicTables")
                    mGraphicTables = True
                    
                Case LCase("frmGoodTurnOver")
                    mfrmGoodTurnOver = True
                    
                Case LCase("frmCreditCustomerAccount")
                    mfrmCreditCustomerAccount = True
                    
                Case LCase("ChangeTotalRemainingAmount")
                    mChangeTotalRemainingAmount = True
                
                Case LCase("AccfrmAsnad")
                    mAccfrmAsnad = True
                
                Case LCase("AccfrmCheckBook")
                    mAccfrmCheckBook = True
                    
                Case LCase("SalarySystem")
                    mSalarySystem = True
                    
                Case LCase("AccfrmKartHesabReport")
                    mAccfrmKartHesabReport = True
                    
                Case LCase("frmTarazSoodZian")
                    mfrmTarazSoodZian = True
                    
                Case LCase("FullSummaryReportSaved")
                    mFullSummaryReportSaved = True
                    
                Case LCase("frmBenefit")
                    mfrmBenefit = True
                    
                Case LCase("frmTempFich")
                    mfrmTempFich = True
                    
                Case LCase("frmHavaleh")
                    mfrmHavaleh = True
                    
                Case LCase("frmTempReceived")
                    mfrmTempReceived = True
                    
                Case LCase("frmSavePurchase")
                    mfrmSavePurchase = True
                    
                Case LCase("frmLosses")
                    mfrmLosses = True
                    
                Case LCase("frmSaveTempReceived")
                    mfrmSaveTempReceived = True
                    
                Case LCase("FirstMojodiControl")
                    mFirstMojodiControl = True
                    
                Case LCase("frmOrder")
                    mfrmOrder = True
                    
                Case LCase("PriceFactorDecrease")
                    mPriceFactorDecrease = True
                    
                Case LCase("SaleReturn")
                    mSaleReturn = True
                    
            End Select
        End If
        Rst.MoveNext
    
    Wend
    
    Set Rst = Nothing
    Set ClsCnn = Nothing
    
End Sub

Private Sub ClearAccess()
    mfrmInvoice = False
    
    mfrmPurchase = False

    mfrmGroupAccounting = False
    
    mfrmGroupBoxTo = False
    
    mfrmPayk = False
    
    mfrmPaykPayment = False
    
    mfrmGroupPer = False
    
    mfrmGroupGoods = False
    
    mfrmGroupSystem = False
    
    mfrmGroupUtility = False
    
    mfrmReports = False
    
    mfrmGroupView = False
    
    mfrmCodingGood = False

    mfrmBackup = False

    mfrmSms = False
    
    mfrmDeviceSetting = False
    
    mfrmCust = False
    
    mfrmCustComp = False
    
    mfrmCustomerDisplay = False
                    
    mFrmDrawerModel = False
    
    mfrmFacEdit = False
    
    mfrmFacRecursive = False
    
    mfrmFind = False
    
    mfrmGood = False
    
    mfrmKB = False
    
    mfrmKeyBoardTZ1 = False
    
    mFrmLogin = False
    
    mfrmMenu = False
    
    mfrmMenuGoodFirst = False
    
    mfrmNotice = False
    
   mfrmStationSetting = False
    
    mfrmPer = False
    
    mfrmPrinter = False
    
    mfrmRep = False
    
    mfrmRestore = False
    
    mfrmSupplier = False
    
    mfrmUsePercent = False
    
    mfrmSupplierComp = False
    
    mfrmWorkTime = False
    
    mmdifrm = False
    
    mfrmSeller = False
    
    mfrmGarson = False
    
    mfrmTable = False

    mfrmGoodPrintFormats = False
    
    mfrmGoodDifferences = False
    
    mEditInvoice = False
    
    mEditPurchase = False
    
    mRefferInvoice = False
    
    mRefferPurchase = False

    mfrmBank = False

    mfrmSalMali = False

    mfrmTagheirSalMali = False

    mfrmFactorReceived = False

    mfrmCredit = False
    
    mfrmCreditCustomer = False
    
    mFrmAccessLevel = False
    
    mfrmPocketPcGroupsAndStations = False
    
    mfrmPocketPCGroupGood = False
    
    mcarryfee = False
    
    mdiscount = False
    
    mEditAllStationsFactors = False
    
    mpacking = False
    
    mRefferedAllStationsFactors = False
    
    mservice = False
    
    mviewallstationsfactors = False

   mfrmHistory = False

    mfrmLablePrinter = False
    
    mfrmGroupStore = False
    
    mfrmBranch = False

    mfrmSupplierComp = False
    
    mfrmSupplier = False
    
    mPrinting = False
    
    mReprint = False
    
    mInvoicePrint = False
    
    mDailyReport = False
    
    mCustomizeFee = False
    
    mfrmCustomReport = False
    
    mfrmCreateSanad = False
    
    mfrmGoodInKitchen = False

    mfrmMojodiControl_1 = False
    
    mfrmMojodiControl_2 = False
    
    mfrmMojodiControl_3 = False
    
    mfrmMojodiControl_4 = False
    
    mfrmRemainingControl = False
    
    mRepGroupGoodsSale = False
    
    mRepTimePerSell = False
    
    mRepDetailGoodsSale = False
    
    mRepBedeh = False
    
    mRepBedehDetail = False
    
    mRepFichSale = False
    
    mRepServeKindSell = False
    
    mRepCustPricDiscount = False
    
    mRepCustPrice = False
    
    mRepSystemGroup = False
    
    mRepSerialFich = False
    
    mRepShift = False
    
    mRepFichBuy = False
    
    mRepDetailGoodsBuy = False
    
    mRepGroupGoodsBuy = False
    
    mRepMojodi = False
    
    mRepGoodList = False
    
    mRepGarson = False
    
    mRepCashInvoice = False
    
    mRepStationSaleSummery = False
    
    mRepCustomerList = False
    
    mRepInventoryRecipt = False
    
    mRepGoodDifferences = False
    
    mRepUsedGoodAmount = False
    
    mRepLossGoodAmount = False
    
    mRepStationSaleSummaryByUser = False
    
    mRepGetOrderGoodAmount = False

    mfrmPayment = False
    
    mMultiPrice = False
    
    mfrmRecieved = False
    
    mfrmAccCoding = False
    
    mfrmCashClose = False
    
    mfrmGoodDiscount = False
    
    mfrmCreditSupplier = False
    
    mRepDailyWeeding = False
    
    mRepDailyPrize = False
    
    mReturnFromAccountDeliver = False
    
    mRepSaleShiftDailyPrize = False
    
    mfrmPhoneBook = False
    
    mEditAllFichUser = False
    
    mRefferedAllFichUser = False

    mfrmDistance = False

    mRepTableSellDetail = False
    
    mfrmManageStore = False

    mfrmCycleStock = False
    
   mfrmWeeding = False
     
   mfrmGoodLevelAutoDiscount = False

    mfrmPrize = False
    
    mfrmPrizeType = False
    
    mfrmInventory = False
    
    mfrmInventory_Level1 = False
    
    mfrmStation_Inventory = False

    mfrmDoreAuction = False
 
    mfrmLevel2Discount = False

    mfrmDoreAuctionRuleDiscount = False

    mfrmColorSizeBrandSeason = False

    mfrmLottery = False

    mfrmDoreLottery = False

    mfrmAutoDiscount = False
    
    mfrmGoodPer = False

    mRepSeller = False
    
    mfrmManageSet = False

    mfrmCustGoodTurnOver = False
    
    mfrmGoodPicture = False

    mEditInvoiceCashClose = False

    mMojodiControl = False
    
    mfrmAccount = False

    mUpperAmountGood = False

    mSaveWithoutPrint = False

    mCashRemaining = False

    mNavigateFactor = False

    mfrmFindFactor = False

    mfrmTaxAndDuty = False

    mfrmTestPos = False

    mLockCheck = False
    
    mReduceAmountGood = False

    mfrmPos = False

    mUnBalance = False

    mGraphicTables = False

    mfrmGoodTurnOver = False
    
    mfrmCreditCustomerAccount = False

    mChangeTotalRemainingAmount = False

    mAccfrmAsnad = False

    mAccfrmCheckBook = False
    
    mSalarySystem = False

    mAccfrmKartHesabReport = False

    mfrmTarazSoodZian = False
    
    mFullSummaryReportSaved = False
    
    mfrmBenefit = False

    mfrmHavaleh = False
    
    mfrmTempReceived = False
    
    mfrmSavePurchase = False
    
    mfrmLosses = False
    
    mfrmSaveTempReceived = False
    
    mFirstMojodiControl = False
    
    mfrmOrder = False
    
    mPriceFactorDecrease = False
    
    mSaleReturn = False
    
    For Each varForm In Forms
        If varForm.Name <> "mdifrm" And varForm.Name <> "frmAbout" Then
            Unload varForm
        End If
    Next

End Sub

