VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicAryaValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 
 ''

Private m_AppPath As String
Private m_dbName As String
Private m_Company As String
Private m_ServerName As String
Private m_StationName As String
Private m_LatinCompany As String
Private m_HardLockSerialNo As String
Private m_ExternalDataName As String
Private m_ExternalDbPath As String
Private m_AccdataBaseName As String
Private m_CustomerDisplayName As String
Private m_CustomerAddres As String
Private m_CustomerTel As String
Private m_CustomerMobile As String
Private m_AccSrvName As String
Private m_AccountSystemName As String
Private m_CustomerString As String

Private m_CustomerId As Integer
Private m_StationNo As Integer
Private m_MaxPrinterNo As Integer
Private m_MaxStationNo As Integer
Private m_MaxPocketPcNo As Integer
Private m_MaxKitchenNo As Integer
Private m_MaxTabletNo As Integer
Private m_MaxAccountingNo As Integer
Private m_EnableUpperAmountGood As Integer
Private m_SecVersion As Integer
Private m_HVersion As Integer
Private m_MiladiDate As Integer

Private m_DemoVersion As Boolean
Private m_SoftLock As Boolean
Private m_HardLock As Boolean
Private m_TrialVer As Boolean
Private m_ProductSystem As Boolean
Private m_LimitedVersion As Boolean
Private m_ExternalDataBase As Boolean
Private m_Accounting As Boolean
Private m_Delivery As Boolean
Private m_BarcodePrint As Boolean
Private m_Customers As Boolean
Private m_ExternalAccounting As Boolean
Private m_StoreGroup As Boolean
Private m_TableGarson As Boolean
Private m_FastPrint As Boolean
Private m_MojodiControl As Boolean
Public SetupFile As String
Private m_MultiPrice As Boolean
Private m_SurroundPayk As Boolean
Private m_GarsonAccess As Boolean
Private m_EditedFichNo As Boolean
Private m_FactorPrintNo As Boolean
Private m_PhoneBook As Boolean
Private m_DifferencePrice As Boolean
Private m_PPCMultiPrice As Boolean
Private m_AdminEdit As Boolean
Private m_BranchView As Boolean
Private m_PrintServer As Boolean
Private m_DBLogin As String
Private m_NetLock As Boolean
Private m_DisableIntro As Boolean
Private m_UnitPrice As String
Private m_DBPass As String
Private m_NewPrinting As Boolean

Public Property Get CustomerId() As Integer
    CustomerId = m_CustomerId
End Property
Public Property Let CustomerId(ByVal vData As Integer)
   m_CustomerId = vData
End Property

Public Property Get LimitedVersion() As Boolean
    LimitedVersion = m_LimitedVersion
End Property
Public Property Let LimitedVersion(ByVal vData As Boolean)
    m_LimitedVersion = LimitedVersion   ' From Database
End Property

Public Property Get HardLockSerialNo() As String
    HardLockSerialNo = m_HardLockSerialNo
End Property
Public Property Let HardLockSerialNo(ByVal vData As String)
   m_HardLockSerialNo = vData
End Property

Public Property Get DemoVersion() As Boolean
    DemoVersion = m_DemoVersion
End Property
Public Property Let DemoVersion(ByVal vData As Boolean)
   m_DemoVersion = vData
End Property
Public Property Get SoftLock() As Boolean
    SoftLock = m_SoftLock
End Property
Public Property Let SoftLock(ByVal vData As Boolean)
   m_SoftLock = vData
End Property
Public Property Get HardLock() As Boolean
    HardLock = m_HardLock
End Property
Public Property Let HardLock(ByVal vData As Boolean)
   m_HardLock = vData
End Property
Public Property Get TrialVer() As Boolean
    TrialVer = m_TrialVer
End Property
Public Property Let TrialVer(ByVal vData As Boolean)
   m_TrialVer = vData
End Property
Public Property Get MaxAccountingNo() As Integer
    MaxAccountingNo = m_MaxAccountingNo
End Property
Public Property Get MaxTabletNo() As Integer
    MaxTabletNo = m_MaxTabletNo
End Property
Public Property Get MaxKitchenNo() As Integer
    MaxKitchenNo = m_MaxKitchenNo
End Property
Public Property Get MaxPocketPcNo() As Integer
    MaxPocketPcNo = m_MaxPocketPcNo
End Property
Public Property Let MaxPocketPcNo(ByVal vData As Integer)
   m_MaxPocketPcNo = vData
End Property
Public Property Get MaxprinterNo() As Integer
    MaxprinterNo = m_MaxPrinterNo
End Property
Public Property Let MaxprinterNo(ByVal vData As Integer)
   m_MaxPrinterNo = vData
End Property
Public Property Get MaxStationNo() As Integer
    MaxStationNo = m_MaxStationNo
End Property
Public Property Let MaxStationNo(ByVal vData As Integer)
   m_MaxStationNo = vData
End Property
Public Property Get DbName() As String
    DbName = m_dbName
End Property
Public Property Let DbName(ByVal vData As String)
    m_dbName = vData   ' From Database
End Property
Public Property Get AccdataBaseName() As String
    AccdataBaseName = m_AccdataBaseName
End Property
Public Property Let AccdataBaseName(ByVal vData As String)
    m_AccdataBaseName = vData   ' From Database
End Property
Public Property Get ServerName() As String
    ServerName = m_ServerName
End Property
Public Property Let ServerName(ByVal vData As String)
   m_ServerName = vData
End Property
Public Property Get AccSrvName() As String
    AccSrvName = m_AccSrvName
End Property
Public Property Let AccSrvName(ByVal vData As String)
   m_AccSrvName = vData
End Property
Public Property Get AccountSystemName() As String
    AccountSystemName = m_AccountSystemName
End Property
Public Property Let AccountSystemName(ByVal vData As String)
   m_AccountSystemName = vData
End Property
Public Property Get Company() As String
    Company = m_Company
End Property
Public Property Let Company(ByVal vData As String)
   m_Company = vData
End Property
Public Property Get LatinCompany() As String
    LatinCompany = m_LatinCompany
End Property
Public Property Get StationName() As String
    StationName = m_StationName
End Property
Public Property Get StationNo() As Integer
    StationNo = m_StationNo
End Property
Public Property Let StationNo(ByVal vData As Integer)
   m_StationNo = vData
End Property
Public Property Get ProductSystem() As Boolean
    ProductSystem = m_ProductSystem
End Property
Public Property Get ExternalDataBase() As Boolean
    ExternalDataBase = m_ExternalDataBase
End Property
Public Property Get ExternalDataName() As String
    ExternalDataName = m_ExternalDataName
End Property
Public Property Let ExternalDataName(ByVal vData As String)
    m_ExternalDataName = vData   ' From Database
End Property
Public Property Get ExternalDbPath() As String
    ExternalDbPath = m_ExternalDbPath
End Property
Public Property Let ExternalDbPath(ByVal vData As String)
    m_ExternalDbPath = vData   ' From Database
End Property
Public Property Get Accounting() As Boolean
    Accounting = m_Accounting
End Property
Public Property Get CustomerDisplayName() As String
    CustomerDisplayName = m_CustomerDisplayName
End Property
Public Property Get Delivery() As Boolean
    Delivery = m_Delivery
End Property
Public Property Get BarcodePrint() As Boolean
    BarcodePrint = m_BarcodePrint
End Property
Public Property Get Customers() As Boolean
    Customers = m_Customers
End Property
Public Property Get ExternalAccounting() As Boolean
    ExternalAccounting = m_ExternalAccounting
End Property
Public Property Let ExternalAccounting(ByVal vData As Boolean)
    m_ExternalAccounting = vData   ' From Database
End Property
Public Property Get StoreGroup() As Boolean
    StoreGroup = m_StoreGroup
End Property
Public Property Get TableGarson() As Boolean
    TableGarson = m_TableGarson
End Property

Public Property Get FastPrint() As Boolean
    FastPrint = m_FastPrint
End Property
Public Property Get MojodiControl() As Boolean
    MojodiControl = m_MojodiControl
End Property
Public Property Get CustomerAddres() As String
    CustomerAddres = m_CustomerAddres
End Property
Public Property Let CustomerAddres(ByVal vData As String)
   m_CustomerAddres = vData
End Property
Public Property Get CustomerTel() As String
    CustomerTel = m_CustomerTel
End Property
Public Property Let CustomerTel(ByVal vData As String)
   m_CustomerTel = vData
End Property
Public Property Get CustomerMobile() As String
    CustomerMobile = m_CustomerMobile
End Property
Public Property Get EnableUpperAmountGood() As Integer
    EnableUpperAmountGood = m_EnableUpperAmountGood
End Property
Public Property Get CustomerString() As String
    CustomerString = m_CustomerString
End Property
Public Property Get MultiPrice() As Boolean
    MultiPrice = m_MultiPrice
End Property
Public Property Get SurroundPayk() As Boolean
    SurroundPayk = m_SurroundPayk
End Property
Public Property Get GarsonAccess() As Boolean
    GarsonAccess = m_GarsonAccess
End Property
Public Property Get EditedFichNo() As Boolean
    EditedFichNo = m_EditedFichNo
End Property
Public Property Get FactorPrintNo() As Boolean
    FactorPrintNo = m_FactorPrintNo
End Property
Public Property Get PhoneBook() As Boolean
    PhoneBook = m_PhoneBook
End Property
Public Property Get DifferencePrice() As Boolean
    DifferencePrice = m_DifferencePrice
End Property
Public Property Get PPCMultiPrice() As Boolean
    PPCMultiPrice = m_PPCMultiPrice
End Property
Public Property Get AdminEdit() As Boolean
    AdminEdit = m_AdminEdit
End Property
Public Property Get BranchView() As Boolean
    BranchView = m_BranchView
End Property
Public Property Get PrintServer() As Boolean
    PrintServer = m_PrintServer
End Property
Public Property Let PrintServer(ByVal vData As Boolean)
   m_PrintServer = vData
End Property
Public Property Get SecVersion() As Integer
    SecVersion = m_SecVersion
End Property
Public Property Get HVersion() As Integer
    HVersion = m_HVersion
End Property
Public Property Get MiladiDate() As Integer
    MiladiDate = m_MiladiDate
End Property
Public Property Get DBLogin() As String
    DBLogin = m_DBLogin
End Property
Public Property Let DBLogin(ByVal vData As String)
   m_DBLogin = vData
End Property
Public Property Get AppPath() As String
    AppPath = m_AppPath
End Property
Public Property Let AppPath(ByVal vData As String)
   m_AppPath = vData
End Property
Public Property Get NetLock() As Boolean
    NetLock = m_NetLock
End Property
Public Property Get DisableIntro() As Boolean
    DisableIntro = m_DisableIntro
End Property
Public Property Get UnitPrice() As String
    UnitPrice = m_UnitPrice
End Property
Public Property Get DBPass() As String
    DBPass = m_DBPass
End Property
Public Property Let DBPass(ByVal vData As String)
   m_DBPass = vData
End Property
Public Property Get NewPrinting() As Boolean
    NewPrinting = m_NewPrinting
End Property

Public Sub Class_Initialize()
    
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
   
    Dim IsFileExist As Boolean
    
    If AryaSettingFile = "" Then End    'Only  For  Make Exe File
    Set tempstring = filetemp.OpenTextFile(AryaSettingFile, ForReading, False, TristateFalse)
    
    Do While tempstring.AtEndOfLine = False
       Str = tempstring.ReadLine
       LenghStr = InStr(1, Str, "=", vbTextCompare)
       
       If InStr(1, Str, "ServerName", vbTextCompare) Then
          m_ServerName = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "dbName", vbTextCompare) Then
          m_dbName = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "StationName", vbTextCompare) Then
          m_StationName = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "StationNo", vbTextCompare) Then
          m_StationNo = Val(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "ExternalDataName", vbTextCompare) Then
          m_ExternalDataName = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ExternalDbPath", vbTextCompare) Then
          m_ExternalDbPath = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "AccSrvName", vbTextCompare) Then
          m_AccSrvName = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "AccountSystemName", vbTextCompare) Then
          m_AccountSystemName = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "AccdataBaseName", vbTextCompare) Then
          m_AccdataBaseName = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "CustomerDisplayName", vbTextCompare) Then
          m_CustomerDisplayName = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "EnableUpperAmountGood", vbTextCompare) Then
          m_EnableUpperAmountGood = CStr(Mid(Str, LenghStr + 1))
          
        ElseIf InStr(1, Str, "ExternalAccounting", vbTextCompare) Then
          m_ExternalAccounting = CStr(Mid(Str, LenghStr + 1))
       
        ElseIf InStr(1, Str, "BranchView", vbTextCompare) Then
          m_BranchView = CStr(Mid(Str, LenghStr + 1))
     
       ElseIf InStr(1, Str, "SecVersion", vbTextCompare) Then
          m_SecVersion = CStr(Mid(Str, LenghStr + 1))
       
        ElseIf InStr(1, Str, "HVersion", vbTextCompare) Then
          m_HVersion = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "PrintServer", vbTextCompare) Then
          m_PrintServer = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "CustomerName", vbTextCompare) Then
          m_Company = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "CustomerAddress", vbTextCompare) Then
          m_CustomerAddres = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "SoftLock", vbTextCompare) Then
          m_SoftLock = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "MiladiDate", vbTextCompare) Then
          m_MiladiDate = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "DBLogin", vbTextCompare) Then
          m_DBLogin = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "AppPath", vbTextCompare) Then
          m_AppPath = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "NetLock", vbTextCompare) Then
          m_NetLock = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "UnitPrice", vbTextCompare) Then
          m_UnitPrice = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "DBPass", vbTextCompare) Then
          m_DBPass = CStr(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "NewPrinting", vbTextCompare) Then
          m_NewPrinting = CStr(Mid(Str, LenghStr + 1))
       
       End If
    
    Loop
    tempstring.Close
    If m_UnitPrice = "" Then m_UnitPrice = "—Ì«·"
    frmfactor.Hide
    StringExeMaker = frmfactor.Label7.Caption
    'StringExeMaker = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
     
    Call CustInfo     ' Get Customer Setting
    
End Sub
Private Sub CustInfo()
    SecurityVersion = Val(clsArya.SecVersion)
    
    If m_DBPass <> "" Then
        SqlPass = m_DBPass
    Else
        m_DBPass = "lemon7430"
        SqlPass = "lemon7430"
    End If
    m_LimitedVersion = False
    If DebugMode = True Then m_LimitedVersion = False
             
    Select Case DebugMode
        Case True
            
            m_HardLock = False
            m_DemoVersion = True   ' For Debug it Will Be True
            
            m_ExternalDataBase = False
            
            strDelegate = "00"                                              'Mid(StringExeMaker, 1, 2)
            strCategory = "00"                                              'Mid(StringExeMaker, 3, 2)
            m_CustomerId = 0                                           'Val(Mid(StringExeMaker, 5, 32))
            intVersion = gold                                  'Val(Mid(StringExeMaker, 8, 1))
            m_MaxStationNo = 12
            m_MaxPrinterNo = 10                                             'Val(Mid(StringExeMaker, 13, 2))
            m_MaxPocketPcNo = 10                                           'Val(Mid(StringExeMaker, 16, 2))
            m_MaxKitchenNo = 2                                                 'Val(Mid(StringExeMaker, 19, 2))
            m_MaxTabletNo = 20                                            'Val(Mid(StringExeMaker, 22, 2))
            m_MaxAccountingNo = 0                                         'Val(Mid(StringExeMaker, 25, 2))

''           m_HardLockSerialNo = "2006-8504-3634"                               'Mid(StringExeMaker, 28, 20)
''           m_HardLockSerialNo = "2006-8508-3727"                             'Mid(StringExeMaker, 28, 20)
''           m_HardLockSerialNo = "85110400043"                             'Mid(StringExeMaker, 28, 20)
''           m_HardLockSerialNo = "2006-8504-2998"                             'Mid(StringExeMaker, 28, 20)
            
'            If m_HardLock = False Then
                m_Company = " ¬—Ì« "                             'Trim(Mid(StringExeMaker, 49, 50))
                m_CustomerAddres = " Â—«‰ - Œ „ÿÂ—Ì - Œ Å«—”« Å·«ò 103"        'Mid(StringExeMaker, 151, 100)
                m_LatinCompany = "Fgarya Co"                                  'Mid(StringExeMaker, 100, 50)
'                m_Company = Space(120) & "‰„«Ì‰œêÌ ê—ê«‰ "                            'Trim(Mid(StringExeMaker, 49, 50))
'                m_CustomerAddres = "ê—ê«‰ "        'Mid(StringExeMaker, 151, 100)
'                m_LatinCompany = " Gorgan Delegate"                                  'Mid(StringExeMaker, 100, 50)
'            End If

'            m_HardLockSerialNo = "87092601050"                             'Mid(StringExeMaker, 28, 20)
            m_HardLockSerialNo = "851104000001" '"86110700724"                             'Mid(StringExeMaker, 28, 20)
'            m_HardLockSerialNo = "93061701000"   ' "85110400001" '"89111701000"  '"89060101559" '"87012800805" '"86110700724"                             'Mid(StringExeMaker, 28, 20)
                             
             m_FastPrint = True
             m_Accounting = True
             m_Customers = True
             HasRfidReader = True
             HasMiniAcc = True
             HasAlbum = True
             HasExcell = True
             HasCRM = False
             
        Case False
             
            m_HardLock = True
            
            If m_LimitedVersion = True Then
                m_ExternalDataBase = False
                m_HardLock = False
                strDelegate = "00"                                                    'Mid(StringExeMaker, 1, 2)
                strCategory = "00"                                              'Mid(StringExeMaker, 3, 2)
                m_CustomerId = 0                                           'Val(Mid(StringExeMaker, 5, 32))
                intVersion = gold                                       'Val(Mid(StringExeMaker, 8, 1))
                m_MaxStationNo = 1
                m_MaxPrinterNo = 4                                             'Val(Mid(StringExeMaker, 13, 2))
                m_MaxPocketPcNo = 0                                           'Val(Mid(StringExeMaker, 16, 2))
                m_MaxKitchenNo = 0                                                 'Val(Mid(StringExeMaker, 19, 2))
                m_MaxTabletNo = 0                                            'Val(Mid(StringExeMaker, 22, 2))
                m_MaxAccountingNo = 0                                         'Val(Mid(StringExeMaker, 25, 2))
                
                m_LatinCompany = "Fgarya Shopping System"                                  'Mid(StringExeMaker, 100, 50)
            
                m_HardLockSerialNo = "12345678901" '"89111701000"  '"89060101559" '"87012800805" '"86110700724"                             'Mid(StringExeMaker, 28, 20)
                
            Else
               
                m_ExternalDataBase = False
                
''***************New String 350 Character and nvcserial =12 and customerid = xxxxx   ************************************
                strDelegate = Mid(StringExeMaker, 19, 2)
                strCategory = Mid(StringExeMaker, 21, 2)
                m_CustomerId = Val(Mid(StringExeMaker, 23, 5))
                intVersion = Val(Mid(StringExeMaker, 28, 1))
                m_HardLockSerialNo = Trim(Mid(StringExeMaker, 30, 18))

                m_MaxStationNo = Val(Mid(StringExeMaker, 1, 2))
                m_MaxPrinterNo = Val(Mid(StringExeMaker, 4, 2))
                m_MaxPocketPcNo = Val(Mid(StringExeMaker, 7, 2))
                m_MaxKitchenNo = Val(Mid(StringExeMaker, 10, 2))
                m_MaxTabletNo = Val(Mid(StringExeMaker, 13, 2))
                m_MaxAccountingNo = Val(Mid(StringExeMaker, 16, 2))

                m_CustomerTel = Trim(Mid(StringExeMaker, 49, 49))
                m_Company = Space(120) & Trim(Mid(StringExeMaker, 99, 50))
                m_CustomerAddres = Trim(Mid(StringExeMaker, 201, 100))
                m_LatinCompany = Trim(Mid(StringExeMaker, 150, 50))
    
                If m_HardLock = True Then
                    RegistryObectvar = CBool(Val(Mid(StringExeMaker, 302, 1)))   ' If True then exe is 2Lock else Exe is Lock
                Else
                    RegistryObectvar = True
                End If

                HasPcPos = CBool(Val(Mid(StringExeMaker, 304, 1)))
                HasTTMS = CBool(Val(Mid(StringExeMaker, 306, 1)))
                HasAryaSms = CBool(Val(Mid(StringExeMaker, 308, 1)))
                HasRfidReader = CBool(Val(Mid(StringExeMaker, 310, 1)))
                HasMiniAcc = CBool(Val(Mid(StringExeMaker, 312, 1)))
                HasAlbum = CBool(Val(Mid(StringExeMaker, 314, 1)))
                HasExcell = CBool(Val(Mid(StringExeMaker, 316, 1)))
                HasCRM = CBool(Val(Mid(StringExeMaker, 318, 1)))
        
''***************Old String 300 Character and nvcserial =11 and customerid = xxx ************************************

'                m_MaxStationNo = Val(Mid(StringExeMaker, 1, 2))
'                m_MaxPrinterNo = Val(Mid(StringExeMaker, 4, 2))
'                m_MaxPocketPcNo = Val(Mid(StringExeMaker, 7, 2))
'                m_MaxKitchenNo = Val(Mid(StringExeMaker, 10, 2))
'                m_MaxTabletNo = Val(Mid(StringExeMaker, 13, 2))
'                m_MaxAccountingNo = Val(Mid(StringExeMaker, 16, 2))

'                strDelegate = Mid(StringExeMaker, 19, 2)
'                strCategory = Mid(StringExeMaker, 21, 2)
'                m_CustomerId = Val(Mid(StringExeMaker, 23, 3))
'                intVersion = Val(Mid(StringExeMaker, 26, 1))
'                m_HardLockSerialNo = Trim(Mid(StringExeMaker, 28, 20))
'
'                m_CustomerTel = Trim(Mid(StringExeMaker, 49, 49))
'                m_Company = Space(120) & Trim(Mid(StringExeMaker, 99, 50))
'                m_CustomerAddres = Trim(Mid(StringExeMaker, 201, 100))
'                m_LatinCompany = Trim(Mid(StringExeMaker, 150, 50))
    
    '***************For make customer exe directly*****************************

'                strDelegate = "00"                                              'Mid(StringExeMaker, 1, 2)
'                strCategory = "00"                                              'Mid(StringExeMaker, 3, 2)
'                m_CustomerId = 1101                                            'Val(Mid(StringExeMaker, 5, 32))
'                intVersion = Diamond                                       'Val(Mid(StringExeMaker, 8, 1))
'                m_MaxStationNo = 5
'                m_MaxPrinterNo = 5                                             'Val(Mid(StringExeMaker, 13, 2))
'                m_MaxPocketPcNo = 2                                           'Val(Mid(StringExeMaker, 16, 2))
'                m_MaxKitchenNo = 2                                                 'Val(Mid(StringExeMaker, 19, 2))
'                m_MaxTabletNo = 3                                            'Val(Mid(StringExeMaker, 22, 2))
'                m_MaxAccountingNo = 0                                         'Val(Mid(StringExeMaker, 25, 2))
'                m_Company = Space(120) & "«› ÃÌ ¬—Ì«  "                            'Trim(Mid(StringExeMaker, 49, 50))
'                m_CustomerAddres = "  Â—«‰  "        'Mid(StringExeMaker, 151, 100)
'                m_LatinCompany = " Fgarya "                                  'Mid(StringExeMaker, 100, 50)
'                m_HardLockSerialNo = "940817004442" ' "93061701000"  ' "85110400001" '"89111701000"  '"89060101559" '"87012800805" '"86110700724"                             'Mid(StringExeMaker, 28, 20)
'                If m_HardLock = True Then
'                    RegistryObectvar = False
'                Else
'                    RegistryObectvar = True
'                End If
'                HasPcPos = True
'                HasTTMS = True
'                HasAryaSms = True
'                HasMiniAcc = True
'                HasAlbum = True
'                HasExcell = True
'                HasRfidReader = True
'
'                strDelegate = "00"                                              'Mid(StringExeMaker, 3, 2)
'                strCategory = "01"                                              'Mid(StringExeMaker, 3, 2)
'                m_CustomerId = 432                                           'Val(Mid(StringExeMaker, 5, 32))
'                intVersion = Diamond                                      'Val(Mid(StringExeMaker, 8, 1))
'                m_MaxStationNo = 1
'                m_MaxPrinterNo = 2                                             'Val(Mid(StringExeMaker, 13, 2))
'                m_MaxPocketPcNo = 0                                           'Val(Mid(StringExeMaker, 16, 2))
'                m_MaxKitchenNo = 0                                                 'Val(Mid(StringExeMaker, 19, 2))
'                m_MaxTabletNo = 0                                            'Val(Mid(StringExeMaker, 22, 2))
'                m_MaxAccountingNo = 0                                         'Val(Mid(StringExeMaker, 25, 2))
'                m_Company = Space(120) & "—” Ê—«‰ ·«Ê«‰”Ì  "                            'Trim(Mid(StringExeMaker, 49, 50))
'                m_CustomerAddres = "  Â—«‰ - „—ò“ Œ—Ìœ Å«·«œÌÊ„ "        'Mid(StringExeMaker, 151, 100)
'                m_LatinCompany = " LoveOnSea "                                  'Mid(StringExeMaker, 100, 50)
'                m_HardLockSerialNo = "93032703304" ' "93061701000"  ' "85110400001" '"89111701000"  '"89060101559" '"87012800805" '"86110700724"                             'Mid(StringExeMaker, 28, 20)
                
'                strDelegate = "00"                                              'Mid(StringExeMaker, 3, 2)
'                strCategory = "01"                                              'Mid(StringExeMaker, 3, 2)
'                m_CustomerId = 430                                           'Val(Mid(StringExeMaker, 5, 32))
'                intVersion = Diamond                                      'Val(Mid(StringExeMaker, 8, 1))
'                m_MaxStationNo = 4
'                m_MaxPrinterNo = 6                                             'Val(Mid(StringExeMaker, 13, 2))
'                m_MaxPocketPcNo = 0                                           'Val(Mid(StringExeMaker, 16, 2))
'                m_MaxKitchenNo = 0                                                 'Val(Mid(StringExeMaker, 19, 2))
'                m_MaxTabletNo = 0                                            'Val(Mid(StringExeMaker, 22, 2))
'                m_MaxAccountingNo = 0                                         'Val(Mid(StringExeMaker, 25, 2))
'                m_Company = Space(120) & "œ› —„—ò“Ì ·«Ê«‰”Ì  "                            'Trim(Mid(StringExeMaker, 49, 50))
'                m_CustomerAddres = "  Â—«‰ - ŒÌ«»«‰ „Ì—⁄„«œ "        'Mid(StringExeMaker, 151, 100)
'                m_LatinCompany = " LoveOnSea_Central "                                  'Mid(StringExeMaker, 100, 50)
'                m_HardLockSerialNo = "93032703301" ' "93061701000"  ' "85110400001" '"89111701000"  '"89060101559" '"87012800805" '"86110700724"                             'Mid(StringExeMaker, 28, 20)
                
'                strDelegate = "00"                                              'Mid(StringExeMaker, 1, 2)
'                strCategory = "00"                                              'Mid(StringExeMaker, 3, 2)
'                m_CustomerId = 918                                           'Val(Mid(StringExeMaker, 5, 32))
'                intVersion = Diamond                                       'Val(Mid(StringExeMaker, 8, 1))
'                m_MaxStationNo = 7
'                m_MaxPrinterNo = 12                                             'Val(Mid(StringExeMaker, 13, 2))
'                m_MaxPocketPcNo = 0                                           'Val(Mid(StringExeMaker, 16, 2))
'                m_MaxKitchenNo = 0                                                 'Val(Mid(StringExeMaker, 19, 2))
'                m_MaxTabletNo = 0                                            'Val(Mid(StringExeMaker, 22, 2))
'                m_MaxAccountingNo = 0                                         'Val(Mid(StringExeMaker, 25, 2))
'                m_Company = Space(120) & "„Ã„Ê⁄Â «” ÅÌ ÌÊ "                            'Trim(Mid(StringExeMaker, 49, 50))
'                m_CustomerAddres = " Â—«‰ - „—ò“ Œ—Ìœ Å«·«œÌÊ„  "        'Mid(StringExeMaker, 151, 100)
'                m_LatinCompany = "SPU"                                  'Mid(StringExeMaker, 100, 50)
'                m_HardLockSerialNo = "93090603618" '"93061701000"   ' "85110400001" '"89111701000"  '"89060101559" '"87012800805" '"86110700724"                             'Mid(StringExeMaker, 28, 20)
'
'                strDelegate = "01"                                              'Mid(StringExeMaker, 1, 2)
'                strCategory = "00"                                              'Mid(StringExeMaker, 3, 2)
'                m_CustomerId = 745                                           'Val(Mid(StringExeMaker, 5, 32))
'                intVersion = Diamond                                       'Val(Mid(StringExeMaker, 8, 1))
'                m_MaxStationNo = 7
'                m_MaxPrinterNo = 15                                             'Val(Mid(StringExeMaker, 13, 2))
'                m_MaxPocketPcNo = 0                                           'Val(Mid(StringExeMaker, 16, 2))
'                m_MaxKitchenNo = 0                                                 'Val(Mid(StringExeMaker, 19, 2))
'                m_MaxTabletNo = 5                                            'Val(Mid(StringExeMaker, 22, 2))
'                m_MaxAccountingNo = 2                                         'Val(Mid(StringExeMaker, 25, 2))
'                m_Company = Space(120) & "»«‘ê«Â »Ì «„ œ»·ÌÊ "                            'Trim(Mid(StringExeMaker, 49, 50))
'                m_CustomerAddres = "„‘Âœ- ŒÌ«»«‰ ”Ã«œ ‰»‘ çÂ«——«Â ê·—Ì“  "        'Mid(StringExeMaker, 151, 100)
'                m_LatinCompany = "BMW_Club"                                  'Mid(StringExeMaker, 100, 50)
'                m_HardLockSerialNo = "93083003601"                            'Mid(StringExeMaker, 28, 20)

''                    m_Company = Space(120) & "„Ã „⁄ Å«·«œÌÊ„ "                            'Trim(Mid(StringExeMaker, 49, 50))
''                    m_CustomerAddres = "  Â—«‰ "        'Mid(StringExeMaker, 151, 100)
''                    m_LatinCompany = "Paladium"                                  'Mid(StringExeMaker, 100, 50)
''                     m_HardLockSerialNo = "93061701000"                            'Mid(StringExeMaker, 28, 20)
''                    m_Company = Space(120) & "ê—ÊÂ Â · Â«Ì Â„« "                            'Trim(Mid(StringExeMaker, 49, 50))
''                    m_CustomerAddres = "  Â—«‰ "        'Mid(StringExeMaker, 151, 100)
''                    m_LatinCompany = "  Homa Hotel"                                  'Mid(StringExeMaker, 100, 50)
''                    m_Company = Space(120) & "«› ÃÌ ¬—Ì«  "                            'Trim(Mid(StringExeMaker, 49, 50))
''                    m_CustomerAddres = " „‘Âœ  "        'Mid(StringExeMaker, 151, 100)
''                    m_LatinCompany = " Fgarya "                                  'Mid(StringExeMaker, 100, 50)
''''
           End If
            
    End Select

    SanadCountingRecord = 5000
    
    If DebugMode = False Then
        m_SoftLock = True
'        If m_SoftLock = False Then
'           m_TrialVer = False
'        Else
           m_TrialVer = True
'        End If
    End If
    
    Call VersionDefine
End Sub

Public Sub VersionDefine()
    If (Val(strCategory) >= 0 And Val(strCategory) <= 6) Or Val(strCategory) = 28 Or Val(strCategory) = 30 Then
        mvarCategory = Restaurant
    ElseIf Val(strCategory) = 7 Then
        mvarCategory = Club
    End If
    
    Select Case intVersion
    
        Case EnumVersion.Min
            
            m_FastPrint = True
            m_Accounting = False
            
            m_Customers = True
            If (Val(strCategory) >= 0 And Val(strCategory) <= 7) Or strCategory = "28" Or strCategory = "30" Then
                m_Delivery = True
            Else
                 m_Delivery = False
            End If
            
            m_TableGarson = False
            m_MojodiControl = False
            m_StoreGroup = False
            
            m_ProductSystem = False
            
            m_BarcodePrint = False
            m_MultiPrice = False
            m_SurroundPayk = False
            m_GarsonAccess = False
            m_EditedFichNo = False
            m_FactorPrintNo = False
            m_PhoneBook = False
            m_DifferencePrice = False
            m_PPCMultiPrice = False
            m_AdminEdit = False
        
        Case EnumVersion.Normal
            m_FastPrint = True
            m_Accounting = False
            
            If (Val(strCategory) >= 0 And Val(strCategory) <= 7) Or strCategory = "28" Or strCategory = "30" Then
                m_Delivery = True
            Else
                 m_Delivery = False
            End If
            m_Customers = True
            If Val(strCategory) >= 0 And Val(strCategory) <= 7 Then
                m_TableGarson = True
            Else
                 m_TableGarson = False
            End If
            m_MojodiControl = False
            m_StoreGroup = False
            
            m_ProductSystem = False
            
            m_BarcodePrint = False
            m_MultiPrice = False
            m_GarsonAccess = False
            m_SurroundPayk = True
            m_EditedFichNo = False
            m_FactorPrintNo = False
            m_PhoneBook = False
            m_DifferencePrice = True
            m_PPCMultiPrice = False
            m_AdminEdit = False
            
        Case EnumVersion.Silver
            m_FastPrint = True
            m_Accounting = True
            m_Customers = True
            
            If Val(strCategory) >= 0 And Val(strCategory) <= 7 Then
                m_TableGarson = True
            Else
                 m_TableGarson = False
            End If
           
            m_MojodiControl = True
            m_StoreGroup = True
            
            m_ProductSystem = True
            
            If (Val(strCategory) >= 0 And Val(strCategory) <= 7) Or strCategory = "9" Or strCategory = "10" Or strCategory = "11" Or strCategory = "25" Or strCategory = "28" Or strCategory = "30" Or strCategory = "38" Then
                m_Delivery = True
            Else
                 m_Delivery = False
            End If
            m_BarcodePrint = True
            m_MultiPrice = True
            m_GarsonAccess = False
            m_SurroundPayk = True
            m_EditedFichNo = True
            m_FactorPrintNo = True
            m_PhoneBook = True
            m_DifferencePrice = True
            m_PPCMultiPrice = False
            m_AdminEdit = True
             
        Case EnumVersion.gold, EnumVersion.Diamond
            m_FastPrint = True
            m_Accounting = True
            m_Customers = True
            
            If Val(strCategory) >= 0 And Val(strCategory) <= 7 Then
                m_TableGarson = True
            Else
                 m_TableGarson = False
            End If
           
            m_MojodiControl = True
            m_StoreGroup = True
            
            m_ProductSystem = True
            
            If (Val(strCategory) >= 0 And Val(strCategory) <= 7) Or strCategory = "9" Or strCategory = "10" Or strCategory = "11" Or strCategory = "25" Or strCategory = "28" Or strCategory = "30" Or strCategory = "38" Then
                m_Delivery = True
            Else
                 m_Delivery = False
            End If
            
            m_BarcodePrint = True
            m_MultiPrice = True
            m_GarsonAccess = True
            m_SurroundPayk = True
            m_EditedFichNo = True
            m_FactorPrintNo = True
            m_PhoneBook = True
            m_DifferencePrice = True
            m_PPCMultiPrice = True
            m_AdminEdit = True
    
    End Select


End Sub

