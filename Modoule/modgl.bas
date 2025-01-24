Attribute VB_Name = "modgl"
Option Explicit


Public mvarAnalyzeForm As Boolean
Public HasCRM As Boolean
Public HasExcell As Boolean
Public HasAlbum As Boolean
Public HasMiniAcc As Boolean
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public HasRfidReader As Boolean
Public HasAryaSms As Boolean
Public HasTTMS As Boolean
Public HasPcPos As Boolean
Public Declare Function Beep Lib "kernel32" _
    (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public TempStatus As EnumFactorType
Public RegistryObectvar As Boolean
Public clsDate As New clsDate
Public SairanFlag As Boolean
'Public Accounting As prjAccount.ClsMonitoring  ''For Runing mode
Public Accounting    ' For Debug Mode
Public Tafsili As Long
Public Tafsili_2 As Long
Public Tafsili_3 As Long
Global Refrence_Acc As Long
Public SqlPass As String

Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Public Declare Function GetSystemWow64Directory Lib "kernel32.dll" Alias _
      "GetSystemWow64DirectoryA" (ByVal lpBuffer As String, ByVal uSize As Long) As Integer

Public Const MAX_PATH = 260
Declare Function GetSystemDirectory Lib "kernel32" Alias _
"GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal _
nSize As Long) As Long

Public MahakScaleOCX3
Public Tiny1
Public LoginSucceeded As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_SHOWWINDOW = &H40
Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1

Public NewCallNumber As String
Public PassPhrase As String

Public timeInterval As Long 'A time interval for a lager measure(Minute)
Public AccessAfterClosingcash As Boolean

Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

'The Send Message functions
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLONG Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Execute a program
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal ClassName As String, ByVal classlength As Long) As Long

Private Declare Function PostMessage Lib "user32" _
Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" _
Alias "FindWindowA" (ByVal szClass$, ByVal szTitle$) As Long
Public Const WM_SYSCOMMAND = &H112&
Public Const SC_SCREENSAVE = &HF140&
Public Const WM_CLOSE = &H10
Public Const WM_GETTEXT = &HD
Public Const WM_SETTEXT = &HC
Public IsClosingInvoiceForm As Boolean 'used for Pos Payment
Public strSock_PosRecive As String
Public bolSock_PosIsConnected As Boolean
Public PosTrain1 As String
Public PosTrain2 As String
Public PosTrain3 As String
Public mVarAccessLevelTemp As Integer
Public mvarMemberShipId As Double
Public LastRecordshow As Boolean

Public mvarTable As Integer
Public mvarInvoiceNO As Long

Public Const STRING_DELIMITER As String = ","
Public ClsFormAccess As New ClsFormAccess
Public clsAccounting As New clsPublicAccounting
Public clsGoodMenu As New clsPublicGoodMenu

Public CurrentDateNumber As Long
Public LastDateNumber As Long
Public DebugMode As Boolean
Public TempPerFlag As Boolean
Public flgShowOrderDetail As Boolean
Public OrderNo As Double
Public MainPriceType As Integer
Public IsHelp As Boolean
Public ShamsiDateName As String
Public mvarBranch As Integer
Public SstabIndex As Integer
Public mvarIndexNo As Integer
Public mvarTipAmount As Long
Public MaxPurchaseRows As Integer
Public MaxInvoiceRows As Integer
Public PagerNo As Integer
Public formloadFlag As Boolean
Public RepVer As String
Public mvarCurrentLoggedInUserName As String
Public CurrentBranchName As String
Public PosConnection As New ADODB.Connection
Public mvarPaperType As EnumPaperType

Public Enum mvarServiceStatus
    Tip = 3
End Enum

Public Enum Operations
    SetDefaultServerDataRegister = 1
    SetDataStation = 2
    ValidateSetDataStation = 3
    LogOutStation = 4
    ClearDataStation = 5
    GetCodeRegister = 6
    SetDataExpireDate = 7
    GetDataStation = 8
    PrintReport = 9
End Enum

Public strSockSend As String
Public strSockRecive As String
Public bolSockIsConnected As Boolean
Public bolMainGroup As Boolean

Public CrystallConnection As String
Public mvarBarcodeName As String
Public SupermarketFlag As Boolean
Public Server_Dir As String

Public mvarNumberOfUnit As Long
Public mvarSellPrice As Currency
Public mvarGoodName As String
Public mvarGoodWeight As Double
Public mvarGoodCode As Double
Public mvarUnitGood As Integer
Public mvarDisCount As Double
Public mvarInventoryNo As Integer
Public mvarRate As Integer
Public TmpInventory As Integer
Public mvarUnitDescription As String
Public mvarMojodi As Long
Public mvarDuty As Boolean
Public mvarTax As Boolean

Public Call_Priority As Integer
Public Call_RealNumber As String
Public ModemPriority(1 To 8) As String
Public Call_Number(1 To 8) As String

Public sFactorReceived As String

Public CurrentBranch As Integer
Public NewGoodFlag As Boolean
Public InventoryNo As Integer

Public AccountYear As String
Public SecurityCount As Integer
Public CreditCode As Long
Public MojodiControlFlag As Boolean
Public FindCustFlag As Boolean
Public DetailsString1 As String
Public DetailsString2 As String
Public DetailsString3 As String
Public DetailsString4 As String
Public mvarServePlace As EnumServePlace
Public mvarStatus As EnumFactorType
Public mvarAddeditMode As EnumAddEditMode
Public mvarNo As Double
Public strMainKey As String
Public strConnectionString As String
Public AccstrConnectionString As String

Public mvarDeleteMsg As String
Public mcol, DCol As New Collection

Public OutPutMachIndex As String
Public RepairCountIndex As String
'Public ClsCurUser As New ClsCurUser
Public View As New Collection
Public mvarArrow As Boolean

Public mvarTafsili As Long
Public mvarPPNo As Integer
Public mvarCurUserNo As Integer
Public mVarAccessLevel As Integer
Public mvarCountRePrint As Integer
Public mvarCountInvoicePrint As Integer
Public mvarBtnNum, mvarBtnAsc, mvarBtnTZ1, mvarBtnKeyboard2, mvarBtnIndex As Integer
Public mvarMsgSelect As Integer

Public MachineName As String
Public Station_IsAccounting As Boolean
Public Station_IsServer As Boolean
Public Server_IP As String
Public LableFile As String

'Public tempString As String 'for changing the name of a key

Public AccessNewMode As Boolean
Public clsArya As New clsPublicAryaValue
Public clsStation As New clsPublicStationValue
Public clsInvoiceValue As New clsPublicInvoiceValue

Public mvarcode As Double
Public mvarName As String
Public mvarPublicOrderType As EnumOrderType
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public StringExeMaker As String
Public intVersion As EnumVersion
Public strDelegate As String
Public strCategory As String
Public SecurityVersion As Integer
Public mvarStartRate As Integer
Public mvarCategory As EnumCategory
'------------------
Public Enum EnumCategory

    NormalCategory = 1
    Restaurant = 2
    Shop = 3
    Beauty = 4
    Taavoni = 5
    Club = 6
End Enum

Public Enum EnumFactorSortItems

    Code = 0
    AlphaBetic = 1
    Fee = 2
    InputKey = 3

End Enum
Public Enum EnumDefaultCustSearch

    MembershipId = 0
    Name = 1
    Phone = 2
    address = 3

End Enum
Public Enum EnumDefaultPhoneBookSearch

    LastName = 1
    Tel = 2
    FirstName = 3
    
End Enum
Public Enum BtnVal
    vbOn = 1
    vbOff = 2
End Enum

Public Enum TypeOfObject
    vbtxtbox = 1
    vbcmbbox = 2
End Enum

Public Enum EnumMsgBox
    vbYes = 1
    vbNo = 2
End Enum

Public Enum EnumSetDate
    KeyDown = 1
    KeyPress = 2
End Enum

Public Enum EnumIncharge
    None = 0
    Seller = 2
    Payk = 3
    Garson = 9
End Enum

Public Enum EnumServePlace
    Salon = 1
    Delivery = 2
    Out = 4
    Car = 8
    Table = 16
    Internet = 32
End Enum

Public Enum EnumOrderType
    ByPhone = 1
    inPerson = 2
End Enum

Public Enum EnumAddEditMode
    ViewMode = 1
    AddMode = 2
    EditMode = 4
    ManipulateMode = 8
    RefferedMode = 16
    InvoiceFactor = 32
    Perfrage = 64
    NoAddMode = 128
End Enum

Public Enum EnumMenuEditMode
    
    ViewButton = 1
    CodeToButton = 2
    ExchangeButton = 4
    DeleteButton = 8
    RenameButton = 16
    PictureButton = 32
    DeletePicture = 64
End Enum


Public Enum EnumDirection
    FirstRecord = 0
    PreviousRecord = 1
    NextRecord = 2
    LastRecord = 3
End Enum

Public Enum EnumFactorType
    Purchase = 1
    Invoice = 2
    Losses = 3
    PurchaseReturn = 4
    InvoiceReturn = 5
    fromStore = 6
    toStore = 7
    StandardHavale = 8
    TempRecieved = 9
    Order = 10
End Enum
Public Enum EnumAccountingType
    Sale = 1
    Buy = 2
    SaleReturn = 3
    BuyReturn = 4
    Payment = 5
    Recieved = 6
End Enum

Public Enum EnumTypeBascule
    Pand = 0
    Digi = 1
    Pand_TLP = 2
End Enum

Public Enum EnumGoodType
    All = 0
    forBuy = 1
    forSale = 2
    forBuySale = 3
    Intermediate = 4
End Enum

Public Enum EnumStationType
    Server = 1
    PC = 2
    Kitchen = 4
    PocketPC = 8
    Tablet = 16
    Account = 32
End Enum

Public Enum EnumKeyBoardType
    Rb2 = 0
    Promag = 1
    S1 = 2
End Enum

Public Enum EnumPaymentType
    Expensive = 0
    TempPersonPayment = 1
    VamPersonPayment = 2
    MosaedePersonPayment = 3
    SalaryPersonPayment = 4
    SupplierPayment = 5
    CustomerPayment = 6
    PaykPayment = 7
    CashRemain = 8
End Enum

Public Enum EnumRecieveType
    TempPersonRecieve = 0
    VamPersonRecieve = 1
    MosaedePersonRecieve = 2
    CustomerRecieve = 3
    SupplierRecieve = 4
    PaykRecieve = 5
    CashRemain = 6

End Enum

Public Enum EnumDeviceType
    Bascule = 1
    BasculeControler = 2
    Pos = 3
    CashDrawer = 4
    CustomerDisplay = 5
    CardReader = 6
    Modem = 7
    Pager = 8
End Enum

Public Enum EnumVersion
    Normal = 0
    Silver = 1
    gold = 2
    Min = 3
    Diamond = 4
End Enum


Public Enum EnumObjectType
    TextBox = 0
    ComboBox = 1
    MaskEdBox = 2
End Enum

Public Enum EnumPaperType
    Receipt = 1
    A4 = 2
End Enum

Public Enum EnumPosType
    PersianPos = 1
    PasargadPos = 2
    IranKishPos = 3
    EghtesadNovinpos = 4
    SamanPOS = 5
    Saderatpos = 6
End Enum

Public Enum EnumDevice

    Mahak = 1
    mahakcontroller = 2
    
    Pand = 3
    PandController = 4
    
    TowzinElectric = 5
    TowzinElectricController = 6
    
    Sairan = 7
    SaIranController = 8
    
    Cas = 9
    CasController = 10
    
    Digi = 13
     
    MDS14000 = 14
    MDS11000 = 15
    PandRoad = 16
    LIONE = 17
    Mahak_Serial = 18
    Mahak_Net = 19
    
    ithacaCashDrawer = 21
    SamsungCashDrawer = 22
    AryaCashDrawer = 23
    NCRCashDrawer = 24
    ABSCashDrawer = 25
    ADPCashDrawer = 26
    EpsonCashDrawer = 27
    DigiCashDrawer = 28
    PartnerCashDrawer = 29
    
    GigaCustomerDisplay = 31
    AryaCustomerDisplay = 32
    NcrCustomerDisplay = 33
    ABSCustomerDisplay = 34
    DigiCustomerDisplay = 36
    EpsonCustomerDisplay = 37
    ZonrichCustomerDisplay = 40
    ZonrichCustomerDisplay_ZQ = 41
    StandardEposCustomerDisplay = 42
    HisenseCustomerDisplay = 43
    HisensePersianCustomerDisplay = 44
   
   MagnetCardReader = 35
   BarcodeSlatReader = 38
   BarcodeTimeReader = 39
   
   CallerIdSharing = 60
   CallerIdModem1 = 61
   CallerIdModem2 = 62
   CallerIdInterface1 = 63
   USBCallerID1 = 64
   CallerIdInterface2_AlmP3 = 65
   CallerIdInterface2_AlmP1 = 66
   CallerIdInterface2_AlmP6 = 67
   RFT230 = 68
   SmsCenter = 82
   
   AryaPager = 81
   
   End Enum
Public Enum EnumCustomerPaymentType

    Cheque = 1
    Loan = 2
    
End Enum
'⁄‰Ê«‰ „ﬁœ«— Ì«  ⁄œ«œ ›«ﬂ Ê—
Public mvarFacDCount  As String
'⁄‰Ê«‰ ﬁÌ„  Ê«Õœ ›«ﬂ Ê—
Public mvarFee As String
'--------------
'„ €Ì—Â«Ì „—»Êÿ »Â Ã” ÃÊ
Public mvarMsgIdx As Integer
Public mvarInput As String
'--------------
'„ €Ì—Â«Ì „—»Êÿ »Â ›Ê‰  Ê—‰ê
Public VarActForm As String
Public CRepFlag As String
Public KindKey As Integer
Public KeyIndex As Integer

'this Enum is related to Language
Public Enum EnumLanguage
    Farsi = 0
    English = 1
End Enum

Public Enum EnumAccessStatus
    None = 0
    Edit = 1
    CashClose = 2
    UpperAmountGood = 3
    LockShow = 4
End Enum

Public Enum EnumAccDocumentType
    Editable = 1
    Temporary = 2
    Permanently = 3
    NoDefinition = 4
End Enum

'Public Type tagInitCommonControlsEx
'   lngSize As Long
'   lngICC As Long
'End Type
'Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
'Public Const ICC_USEREX_CLASSES = &H200

Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Public Sub AllButton(mydata As BtnVal, Optional myVar As Boolean)
Dim i As Integer
If mydata = 1 Then
    For i = 0 To mdifrm.Toolbar1.Buttons.Count - 4
        mdifrm.Toolbar1.Buttons(i + 1).Enabled = True
    Next i
ElseIf mydata = 2 Then
    For i = 0 To mdifrm.Toolbar1.Buttons.Count - 4
        mdifrm.Toolbar1.Buttons(i + 1).Enabled = False
    Next i
End If
If myVar Then
    mdifrm.Toolbar1.Buttons(12).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
Else
    mdifrm.Toolbar1.Buttons(12).Enabled = False
    mdifrm.Toolbar1.Buttons(25).Enabled = True
End If
End Sub

Public Sub Main()
    
    InitCommonControls
'    On Error Resume Next
'    Dim iccex As tagInitCommonControlsEx
'    With iccex
'        .lngSize = LenB(iccex)
'        .lngICC = ICC_USEREX_CLASSES
'    End With
'    InitCommonControlsEx iccex
    
    Dim clsDate As New clsDate
    strMainKey = "Total"
    DebugMode = True
    
    If DebugMode = False Then
        If Val(GetSetting(strMainKey, "frmSplash", "DisableSplash")) = 0 Then
            frmSplash.Show
        Else
            FrmLogin.Show
        End If
    Else
        FrmLogin.Show
    End If

'    frmSplash.Refresh
'    Sleep 5000
'    SleepEx 5000, 0
'    frmSplash.Visible = False
'    Unload frmSplash
'    FrmLogin.Show

End Sub

Public Sub ArrangeGauge()
Dim i, WidthButton As Integer
For i = 1 To mdifrm.Toolbar1.Buttons.Count
    WidthButton = WidthButton + mdifrm.Toolbar1.Buttons(i).Width
Next i
End Sub

Public Sub KeyActi(Obj As TypeOfObject, KeyCode As Integer, Shift As Integer, frmact As Form, Optional NotArrowKey As Boolean)
On Error Resume Next
Dim ShiftKey As Integer
ShiftKey = Shift And 7

On Error Resume Next
 
'''' Call NeccesaryFunction
'''' Unload frmfactor
'''' Unload frmInput
'''' Unload frmMsg
Select Case ShiftKey
    Case 0
        Select Case KeyCode
        
           Case vbKeyPageDown     'PageDown
                    If Not (mvarArrow) Then
                        mvarArrow = True
                        If Obj = vbtxtbox And Not (NotArrowKey) Then
                            If Not IsNull(frmact) Then
                                frmact.BeforePreviousKey
                            End If
                            If Not IsNull(frmact) Then
                            
                                frmact.PreviousKey
                            End If
                        End If
                        mvarArrow = False
                    End If

            Case vbKeyPageUp    'PageUp
                If Not (mvarArrow) Then
                    mvarArrow = True
                    If Obj = vbtxtbox And Not (NotArrowKey) Then
                        If Not IsNull(frmact) Then
                            frmact.BeforeNextKey
                        End If
                        If Not IsNull(frmact) Then
                            frmact.NextKey
                        End If
                    End If
                    mvarArrow = False
                End If
            
            Case vbKeyHome 'Home
                If Not (mvarArrow) Then
                    mvarArrow = True
                    If Obj = vbtxtbox And Not (NotArrowKey) Then
                        If Not IsNull(frmact) Then
                            frmact.BeforeFirstKey
                        End If
                        If Not IsNull(frmact) Then
                            frmact.FirstKey
                        End If
                    End If
                    mvarArrow = False
                End If
            
            
            Case vbKeyEnd    'End
                If Not (mvarArrow) Then
                    mvarArrow = True
                    If Obj = vbtxtbox And Not (NotArrowKey) Then
                        If Not IsNull(frmact) Then
                            frmact.BeforeLastKey
                        End If
                        If Not IsNull(frmact) Then
                            frmact.LastKey
                        End If
                    End If
                    mvarArrow = False
                End If
            
            Case vbKeyInsert    'Add Key (Insert)
                    If Not IsNull(frmact) Then
                        frmact.BeforeAdd
                    End If
                    If Not IsNull(frmact) Then
                        frmact.Add
                    End If
            Case vbKeyF3   'Edit Key (F3)
                    If Not IsNull(frmact) Then
                        frmact.Edit
                    End If
            Case 13:    'Enter Key
''''                    If Not IsNull(frmAct) Then  '  And exit_flag = False Then
''''                       frmAct.Update
''''                    End If
                
            Case vbKeyEscape    'Esc Key
''''                    If Not IsNull(frmAct) And exit_flag = False Then
''''                        frmAct.Cancel
''''                    End If
''''
''''                    If Not IsNull(frmAct) Then
''''                        frmAct.AfterCancel
''''                    End If
            Case vbKeyDelete    'Delete Key
            Case vbKeyF2   'Find Key(F2)
                If Not IsNull(frmact) Then
                    frmact.Find
                End If
            Case vbKeyF6   'Printing (F6)
                If Not IsNull(frmact) Then
                    frmact.Printing
                End If
            Case vbKeyF7   'Printing (F6)
            Case vbKeyF9   'Recursive Key (F9)
                If clsStation.KeyboardType <> EnumKeyBoardType.S1 Then
                    If Not IsNull(frmact) Then
                        frmact.UndoRedo
                    End If
                End If
            Case vbKeyF12   'Scan (F12)
'                If Not IsNull(frmAct) Then
'                    frmAct.Scan
'                End If
        End Select
    
   
'    Case vbShiftMask     'Shift Key
               
   Case vbCtrlMask      'Control Key
       
       Select Case KeyCode
           Case vbKeyF8     'Open Drawer
                      
           Case vbKeyF9        'Phone Book in Modgl Routine
                If clsStation.KeyboardType <> EnumKeyBoardType.Promag Then
                    frmPhoneBook.Show
                    frmPhoneBook.SetFocus
                End If
          
            Case vbKeyF12   'Exit Form (Ctl + F12)
                If clsStation.KeyboardType <> EnumKeyBoardType.S1 Then
                    If Not IsNull(frmact) Then
                        frmact.ExitForm
'                        mdifrm.fwBtnCtrl.SetFocus
                    End If
                End If
       End Select
 ' Case 4     ' Alt Key

End Select
Exit Sub
ErrorHandler1:
        MsgBox err.Description
        MsgBox err.Source
        'MsgBox err.
End Sub

Public Sub HeaderLabel(mydata As Integer, Obj As Object)
    Select Case clsStation.Language
    
        Case Farsi
            Select Case mydata
                Case AddMode
                    Obj.Caption = "ÃœÌœ"
                Case EditMode
                    Obj.Caption = "«’·«Õ"
                Case ManipulateMode
                    Obj.Caption = " €ÌÌ—« "
                Case ViewMode
                    Obj.Caption = "„—Ê—"
                Case RefferedMode
                    Obj.Caption = "„—ÃÊ⁄Ì"
            End Select
            
        Case English
        
            Select Case mydata
                Case AddMode
                    Obj.Caption = "New"
                Case EditMode
                    Obj.Caption = "Edit"
                Case ManipulateMode
                    Obj.Caption = "Manipulation"
                Case ViewMode
                    Obj.Caption = "View"
                Case RefferedMode
                    Obj.Caption = "Reffered"
            End Select
            
    End Select
End Sub

Public Sub SetDate(mydata As EnumSetDate, ByRef Obj As TextBox, ByRef Key As Integer)
Select Case mydata
    Case KeyDown:
        If Key = vbKeyDelete Then
            If Mid(Obj.Text, Obj.SelStart + 1, 1) = "/" Then
                Key = 0
            End If
        End If
    Case KeyPress:
        On Error Resume Next
        If Len(Obj.Text) >= 8 And (Key >= 48 And Key <= 57) Then
            Key = 0
            Exit Sub
        End If
        If Key = 8 Then
            If Len(Obj.Text) = Obj.SelStart Then
                Exit Sub
            End If
            If Mid(Obj.Text, Obj.SelStart, 1) = "/" Then
                Key = 0
                Exit Sub
            End If
            Exit Sub
        End If
        If Key < 48 Or Key > 57 Then
            Key = 0
            Exit Sub
        End If
        If Len(Obj.Text) <> Obj.SelStart Then
            Exit Sub
        End If
        Select Case Len(Obj.Text)
        Case 2
            Obj.Text = Obj.Text & "/"
            Obj.SelStart = Len(Obj.Text) + 1
        Case 5
            Obj.Text = Obj.Text & "/"
            Obj.SelStart = Len(Obj.Text) + 1
        End Select
End Select
End Sub

Public Function IsUserDefinedKey(KeyCode As Integer, Shift As Integer)
    Dim ReturnValue As Boolean
    ReturnValue = False

     If KeyCode >= 65 And KeyCode <= 90 Then          '
          ReturnValue = True
     ElseIf KeyCode >= 48 And KeyCode <= 57 And Shift <> 0 Then      'Digits
          ReturnValue = True
     ElseIf KeyCode >= 112 And KeyCode <= 123 And Shift = 1 Then    'Shift +(F1 ~ F12)
          ReturnValue = True
     ElseIf KeyCode >= 128 Then     '
          ReturnValue = True
    End If
    If (KeyCode = 187 Or KeyCode = 189 Or KeyCode = 190 Or KeyCode = 191) And Shift = 0 Then       ' = - . /
          ReturnValue = False
    End If
    If KeyCode = 53 And Shift = 1 Then        '%
           ReturnValue = False
    End If
    If KeyCode = 220 And Shift = 2 Then        'Alarm(Ctrl + \)
           ReturnValue = False
    End If
''''    If (KeyCode = 86 Or KeyCode = 88) And Shift = 2 Then        '(Ctrl + v)& (Ctrl + x)Only For Rb2
''''           ReturnValue = False
''''    End If
    If (KeyCode = 52 Or KeyCode = 222) And Shift = 1 And clsStation.KeyboardType = EnumKeyBoardType.Rb2 Then         '%
           ReturnValue = False
    End If
   
    If clsStation.AlphabeticGoods = True Then
       If KeyCode >= 65 And KeyCode <= 90 And Shift = 0 Then          '
          ReturnValue = False
       End If
       If (KeyCode = 186 Or KeyCode = 188 Or KeyCode = 192 Or KeyCode = 219 Or KeyCode = 220 Or KeyCode = 221 Or KeyCode = 222) And Shift = 0 Then            '
          ReturnValue = False
       End If
       If KeyCode = 72 And Shift = 1 Then          'H(¬)
          ReturnValue = False
       End If
    End If
    IsUserDefinedKey = ReturnValue
End Function
Public Function mvarShiftNo() As Integer
    mvarShiftNo = 0
    Dim cnn As New ADODB.Connection
    Dim Rst As New ADODB.Recordset
    
    cnn.ConnectionString = strConnectionString
    cnn.Open
'    ReDim Parameter(0) As Parameter
'    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
'    Set Rst = RunParametricStoredProcedure2Rec("Get_All_tShift", Parameter, cnn)
    Set Rst = RunStoredProcedure2RecordSet("Get_All_tShift")
    If Not (Rst.EOF = True And Rst.BOF = True) Then
      '  Rst.MoveFirst
        While Rst.EOF <> True
            If Rst.Fields("StartTime") > Rst.Fields("EndTime") Then
                If time >= Rst.Fields("StartTime") Or time < Rst.Fields("EndTime") Then
                    mvarShiftNo = Rst.Fields("Code")
                    
                End If
            Else
                If time >= Rst.Fields("StartTime") And time < Rst.Fields("EndTime") Then
                    mvarShiftNo = Rst.Fields("Code")
                End If
            End If
            Rst.MoveNext
        Wend
    End If
    
    Set Rst = Nothing
    Set cnn = Nothing
    
End Function

Public Function mvarDate() As String

    Dim Parameter() As Parameter
    Dim CnnDate As New ADODB.Connection
    Dim Rst As New ADODB.Recordset
    Dim clsDate As New clsDate

    CnnDate.ConnectionString = strConnectionString
    CnnDate.Open
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, mvarShiftNo)
'    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_tShift_By_Code", Parameter, CnnDate)
    If Rst.EOF <> True And Rst.BOF <> True Then
        If Rst.Fields("starttime") > Rst.Fields("Endtime") And time < Rst.Fields("Endtime") Then
            mvarDate = Trim(Right(clsDate.shamsi(DateAdd("d", -1, Now)), 8))
        Else
            mvarDate = Trim(Right(clsDate.shamsi(Date), 8))
        End If
    Else
        mvarDate = Trim(Right(clsDate.shamsi(Date), 8))
    End If
    Set Rst = Nothing
    Set CnnDate = Nothing

End Function

Public Function DText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim StrBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Decrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            StrBuff = StrBuff & Chr$(c And &HFF)
        Next i
    Else
        StrBuff = strText
    End If
    DText = StrBuff
End Function

Public Function DateToNumber(strDate As Variant) As Long
    Dim y As String
    Dim m As String
    Dim D As String

 '   On Error GoTo ErrorHandle

    If Len(strDate) = 0 Or strDate = "____/__/__" Then
        DateToNumber = 0
        Exit Function
    End If
    y = Mid(strDate, 1, 4)
    m = Mid(strDate, 6, 2)
    If InStr(m, "/") > 0 Then m = "0" + Left(m, 1)
    D = Right(strDate, 2)
    If InStr(D, "/") > 0 Then D = "0" + Right(D, 1)
     DateToNumber = CLng(y + m + D)
End Function
Public Function DateToNumber8(strDate As Variant) As Long
    Dim y As String
    Dim m As String
    Dim D As String

 '   On Error GoTo ErrorHandle

    If Len(strDate) = 0 Or strDate = "__/__/__" Then
        DateToNumber8 = 0
        Exit Function
    End If
    y = Mid(strDate, 1, 2)
    m = Mid(strDate, 4, 2)
    If InStr(m, "/") > 0 Then m = "0" + Left(m, 1)
    D = Right(strDate, 2)
    If InStr(D, "/") > 0 Then D = "0" + Right(D, 1)
    DateToNumber8 = CLng(y + m + D)
End Function

Public Function NumberToDate(lngDate As Variant) As String
    Dim y As String
    Dim m As String
    Dim D As String
    Dim a As Integer

    NumberToDate = "__/__/__"

    If Val(lngDate) <= 0 Then
        NumberToDate = "__/__/__"
        Exit Function
    End If
    a = lngDate / 10000
    y = CStr(a)
    a = (lngDate / 100) Mod 100
    m = CStr(a)
    a = lngDate Mod 100
    D = CStr(a)
    If InStr(D, "/") > 0 Then D = "0" + Right(D, 1)
    NumberToDate = Format$(y, "00") + "/" + Format$(m, "00") + "/" + Format$(D, "00")
End Function

Public Function TestChasban(ByVal c1 As String, ByVal c2 As String)
 Dim SmallCh, Ch As String
 Dim IsSmall, Xh  As Integer
  SmallCh = "ïùìóôõü°®™¨ÆØ‡‰ËÍÏÓÛı˜˚˛"
  IsSmall = 0
  For Xh = 1 To Len(SmallCh)
    If c2 = Mid(SmallCh, Xh, 1) Then IsSmall = 1
  Next
  Ch = c1
  If (IsSmall = 1) And ((c1 = "ê") Or (c1 = "‰") Or (c1 = "Ë") Or (c1 = "˛")) Then
  Select Case c1
    Case "ê":
      Ch = "ë"
    Case "‰":
      Ch = "„"
    Case "Ë":
      Ch = "Á"
    Case "˝":
      Ch = "¸"
  End Select
  End If
  TestChasban = Ch
End Function


Public Function BigChar(Ch As String) As String
  Dim BigCh As String
  BigCh = "ÄÅÇÉÑÖÜáàâäãåçéèëëííîîññòòööúúûû††¢£§•¶ßß©©´´≠≠Ø∞±≤≥¥µ∂∑∏π∫ªºΩæø¿¡¬√ƒ≈∆«»… ÀÃÕŒœ–—“”‘’÷◊ÿŸ⁄€‹›ﬁﬂ‡·‚‚·ÂÊÊÂÈÈÎÎÌÌÔÔÒÚÒÙÙˆˆ¯˘˘˘¸˝¸ˇ"
  If Asc(Ch) > 128 Then
    Ch = Mid(BigCh, Asc(Ch) - 127, 1)
  End If
 BigChar = Ch
End Function

Public Function WinToIranSys(st As String)
 
 Dim DimFars As String
 Dim NumberEng As String
 Dim FarsiChar As String
      FarsiChar = "                                               ÄÅÇÉÑÖÜáàâ       ABCDEFGHIJKLMNOPQRSTUVWXYZ      abcdefghijklmnopqrstuvwxyz     "
      FarsiChar = FarsiChar + " ï           ù¶                                                  ç   èêì óôõü°¢£§•®™¨Æ Ø‡‰Ë ÍÏÓ Û ı˜˚¯     ¸˛                 "
  NumberEng = "01234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
  
  Dim Sh As String
 '   Xh As word
 Dim Ch As String
 Dim Nu, Xh As Integer
  
    Nu = 1
    Sh = ""
   For Xh = 1 To Len(st)
      Ch = Mid(FarsiChar, Asc(Mid(st, Xh, 1)), 1)
      If (InStr(1, NumberEng, Mid(st, Xh, 1), 1)) Or ((Nu > 1) And (Mid(st, Xh, 1) = " ")) Then
        Sh = Left(Sh, Nu - 1) + Ch + Mid(Sh, Nu, Len(Sh))
        Nu = Nu + 1
      Else
        Sh = Ch + Sh
        Nu = 1
      End If
    If (Ch = " ") And (Xh > 1) Then
        Sh = Left(Sh, 1) + BigChar(Mid(Sh, 2, 1)) + Mid(Sh, 3, Len(Sh))
    End If
    Dim S1 As String
    S1 = Mid(Sh, 2, Len(Sh))
    If Xh = Len(st) Then
      Sh = BigChar(Mid(Sh, 1, 1)) + S1
    End If
   
    If Len(Sh) > 1 Then
      S1 = Mid(Sh, 2, Len(Sh))
      Sh = TestChasban(Mid(Sh, 1, 1), Mid(Sh, 2, 1)) + S1
    End If
    If (Ch = " ") And (Xh > 1) And (Mid(Sh, 2, 1) <> "ê") Then
      S1 = BigChar(Mid(Sh, 2, 1))
      Sh = Mid(Sh, 1, 1) + S1 + Mid(Sh, 3, Len(Sh))
    End If
    If (Xh = Len(st)) And (Mid(Sh, 1, 1) <> "ê") Then
      Sh = BigChar(Mid(Sh, 1, 1)) + Mid(Sh, 2, Len(Sh))
    End If
   Next Xh
   WinToIranSys = Sh
   
End Function
Function Decrypt(encryptedString As String) As String
   Dim secureBytes() As Byte
   Dim index As Integer
   
   secureBytes = encryptedString
   
   For index = LBound(secureBytes) To UBound(secureBytes)
        secureBytes(index) = secureBytes(index) Xor 1
   Next index
    
   Decrypt = secureBytes
   MsgBox Decrypt
   
End Function

Function DecryptASCII(encryptedString As String) As String

    Dim secureBytes() As Byte
    Dim index As Integer
    
    secureBytes = encryptedString
    
    For index = LBound(secureBytes) To UBound(secureBytes)
        If secureBytes(index) > 10 Then
            secureBytes(index) = secureBytes(index) - 10
        End If
    Next index
    
    DecryptASCII = secureBytes
    MsgBox DecryptASCII, vbOKOnly, "Decrypted String"
End Function

Public Sub LogSave(ByVal InputString As String, Optional ByRef ErrorObject As ErrObject, Optional ByVal SourceProc As String)
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim CallerIDFile As String
    Dim ErrorString As String
    
    ErrorString = ""
    
    If Not IsNull(ErrorObject) Then
        ErrorString = vbCrLf & Now & vbCrLf & InputString & SourceProc & vbCrLf & "Description=> " & ErrorObject.Description & vbCrLf & "Error Number=> " & ErrorObject.Number & vbCrLf & "Error Source=> " & ErrorObject.Source
        ErrorObject.Clear
    End If
    
   If Not filetemp.FolderExists(App.Path & "\Log") Then
        filetemp.CreateFolder App.Path & "\Log"
   End If
    CallerIDFile = App.Path & "\Log\" & DateToNumber8(Right(clsDate.shamsi(Date), 8)) & ".Log"
'    CallerIDFile = App.Path & ".Log"
    If filetemp.FileExists(CallerIDFile) Then
        Set tempstring = filetemp.OpenTextFile(CallerIDFile, ForAppending, False, TristateFalse)
    Else
        filetemp.CreateTextFile CallerIDFile
        Set tempstring = filetemp.OpenTextFile(CallerIDFile, ForWriting, False, TristateFalse)
    End If
    tempstring.WriteLine (ErrorString)
    tempstring.Close

End Sub

Public Sub LogSaveNew(ByVal InputString As String, Optional ByVal ErrorDescription As String = "", Optional ByVal ErrorNumber As String = "", Optional ByVal ErrorSource As String = "", Optional ByVal SourceProc As String = "")
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim CallerIDFile As String
    Dim ErrorString As String
    
    ErrorString = ""
    
    If ErrorDescription <> "" And ErrorNumber <> "" And ErrorSource <> "" Then
        ErrorString = vbCrLf & Now & vbCrLf & InputString & SourceProc & vbCrLf & "Description=> " & ErrorDescription & vbCrLf & "Error Number=> " & ErrorNumber & vbCrLf & "Error Source=> " & ErrorSource
    Else
        ErrorString = vbCrLf & Now & InputString & vbCrLf
    End If
    
   If Not filetemp.FolderExists(App.Path & "\Log") Then
        filetemp.CreateFolder App.Path & "\Log"
   End If
    CallerIDFile = App.Path & "\Log\" & DateToNumber8(Right(clsDate.shamsi(Date), 8)) & ".Log"
'    CallerIDFile = App.Path & ".Log"
    If filetemp.FileExists(CallerIDFile) Then
        Set tempstring = filetemp.OpenTextFile(CallerIDFile, ForAppending, False, TristateFalse)
    Else
        filetemp.CreateTextFile CallerIDFile
        Set tempstring = filetemp.OpenTextFile(CallerIDFile, ForWriting, False, TristateFalse)
    End If
    tempstring.WriteLine (ErrorString)
    tempstring.Close

End Sub

Public Sub ShowMessage(ByVal LabelCaption As String, ByVal btnOkVisible As Boolean, ByVal btnCancelVisible As Boolean, ByVal btnOkCaption As String, ByVal btnCancelCaption As String)
        frmMsg.fwlblMsg.Caption = LabelCaption
        frmMsg.fwBtn(0).Visible = btnOkVisible
        frmMsg.fwBtn(1).Visible = btnCancelVisible
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = btnOkCaption
        frmMsg.fwBtn(1).Caption = btnCancelCaption
        frmMsg.Show vbModal
End Sub

Public Sub ShowDisMessage(ByVal LabelCaption As String, ByVal TimerInterval As Long)
    frmDisMsg.Timer1.Enabled = False
    frmDisMsg.Timer1.Interval = TimerInterval
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.lblMessage.Caption = LabelCaption
    frmDisMsg.Show vbModal
End Sub
Public Sub ShowDisMessageNoModal(ByVal LabelCaption As String, ByVal TimerInterval As Long)
    frmDisMsg.Timer1.Enabled = False
    frmDisMsg.Timer1.Interval = TimerInterval
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.lblMessage.Caption = LabelCaption
    frmDisMsg.Show
End Sub

Public Sub ShowErrorMessage()
        If strDelegate = "24" Then
            frmMsg.fwlblMsg.Caption = "Œÿ« œ— «‰Ã«„ ⁄„·Ì« . ·ÿ›« »—«Ì —›⁄ «‘ﬂ«·° »« ‘—ﬂ   ﬂÌ‰ «·ﬂ —Ê‰Ìﬂ Å«”«—ê«œ  „«” »êÌ—Ìœ"
        Else
            frmMsg.fwlblMsg.Caption = "Œÿ« œ— «‰Ã«„ ⁄„·Ì« . ·ÿ›« »—«Ì —›⁄ «‘ﬂ«·° »« ‘—ﬂ  ›‰ ¬Ê— ê” — ¬—Ì«  „«” »êÌ—Ìœ"
        End If
        frmMsg.fwBtn(0).Visible = True
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = " «ÌÌœ"
        frmMsg.fwBtn(1).Caption = " "
        frmMsg.Show vbModal
End Sub

Public Sub ShowInputForm(ByVal OptionVisible0 As Boolean, ByVal OptionVisible1 As Boolean, ByVal OptionVisible2 As Boolean, ByVal OptionCaption0 As String, ByVal OptionCaption1 As String, ByVal OptionCaption2 As String, ByVal HeaderCaption As String, ByVal ButtonOkVisible As Boolean, ByVal ButtonCancelVisible As Boolean, ByVal TextBoxVisible As Boolean, ByVal DefaultValue As Integer)
    frmInput.OptionLevel(0).Visible = OptionVisible0
    frmInput.OptionLevel(1).Visible = OptionVisible1
    frmInput.OptionLevel(2).Visible = OptionVisible2
    frmInput.OptionLevel(0).Caption = OptionCaption0
    frmInput.OptionLevel(1).Caption = OptionCaption1
    frmInput.OptionLevel(2).Caption = OptionCaption2
    frmInput.fwlblInput.Caption = HeaderCaption
    frmInput.btnOk.Visible = ButtonOkVisible
    frmInput.btnCancel.Visible = ButtonCancelVisible
    frmInput.txtInput.Visible = TextBoxVisible
    frmInput.Picture1.Visible = True
    If DefaultValue >= 0 And DefaultValue <= 2 Then
        frmInput.OptionLevel(DefaultValue).Value = True
    End If
    frmInput.Show vbModal
End Sub

Public Function GetPerInfo(ByVal UserName As String, ByVal Password As String, ByVal Branch As Long) As ADODB.Recordset
    ReDim Parameters(1) As Parameter
    
    Parameters(0) = GenerateInputParameter("@UserName", adVarChar, 50, UserName)
    Parameters(1) = GenerateInputParameter("@PassWord", adVarChar, 50, Password)
'    Parameters(2) = GenerateInputParameter("@Branch", adInteger, 4, Branch)
    
    Set GetPerInfo = RunParametricStoredProcedure2Rec("GetPerInfo", Parameters)

End Function
Public Sub ShowAccountingForm(ByVal strFormName As String, ByVal WinTitle As String)
    On Error Resume Next
'    Dim Accounting As New prjAccount.ClsMonitoring
    Dim hWnd, retval As Long
'    WinTitle = "Õ”«»œ«—Ì"  ' "Recycle Bin" '<- Title of Window
    hWnd = FindWindow(vbNullString, WinTitle) '
'    retval = PostMessage(hWnd, WM_CLOSE, 0&, 0&) ' Close Window
'    If hWnd = 0 Then
'        Accounting.ShowAccountingForms strFormName, clsArya.DBLogin, clsArya.DbName, clsArya.ServerName, CStr(AccountYear), Trim(clsArya.Company), CurrentBranch, mvarCurUserNo
        Accounting.ShowAccountingForms strFormName
        
'    End If
'    hwnd = FindWindow(vbNullString, WinTitle)  '

End Sub
Public Sub ODBCSetting(ByVal ServerName As String, ByVal DbName As String)
    
    Call mdifrm.FWRegistry1.SetKeyStr(flwRegLocalMachine, "software\odbc\odbc.ini\totaluser", "Server", ServerName)
    Call mdifrm.FWRegistry1.SetKeyStr(flwRegLocalMachine, "software\odbc\odbc.ini\totaluser", "DataBase", DbName)
'Xp Mediacenter
    Call mdifrm.FWRegistry1.SetKeyStr(flwRegCurrentUser, "software\odbc\odbc.ini\totaluser", "Server", ServerName)
    Call mdifrm.FWRegistry1.SetKeyStr(flwRegCurrentUser, "software\odbc\odbc.ini\totaluser", "DataBase", DbName)

End Sub
Public Function UpdateHavaleResid(ByVal InventoryNo As Integer, ByVal AccountYear As Integer, ByVal GoodCode As Integer, ByVal BeforeDate As String, ByVal AfterDate As String) As Boolean
    Dim NumberOfRecords As Long
    UpdateHavaleResid = False
    If GoodCode = 0 Then
        ReDim Parameter(6) As Parameter
        Parameter(0) = GenerateInputParameter("@InVentoryNo", adInteger, 4, InventoryNo)
        Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameter(2) = GenerateInputParameter("@GoodCode", adInteger, 4, GoodCode)
        Parameter(3) = GenerateInputParameter("@Flag", adInteger, 4, 0)  ' select Records
        Parameter(4) = GenerateInputParameter("@BeforeDate", adVarWChar, 8, BeforeDate)
        Parameter(5) = GenerateInputParameter("@AfterDate", adVarWChar, 8, AfterDate)
        Parameter(6) = GenerateOutputParameter("@NumberOfRecords", adInteger, 4)
        
        NumberOfRecords = RunParametricStoredProcedure2String("Update_HavalehResid", Parameter)
        Dim CalculateTime As Integer
        CalculateTime = NumberOfRecords / 900
        If CalculateTime > 1 And GoodCode = 0 Then
            ShowMessage "»Â —Ê“ —”«‰Ì ÕœÊœ " & CalculateTime & "œﬁÌﬁÂ ÿÊ· „Ì ò‘œ. ¬Ì« «ò‰Ê‰ „«Ì· »Â «‰Ã«„ ¬‰ Â” Ìœø", True, True, "»·Ì", "ŒÌ—"
            If mvarMsgIdx = vbNo Then Exit Function
        End If
    End If
    mdifrm.MousePointer = vbHourglass
    ReDim Parameter(6) As Parameter
    Parameter(0) = GenerateInputParameter("@InVentoryNo", adInteger, 4, InventoryNo)
    Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(2) = GenerateInputParameter("@GoodCode", adInteger, 4, GoodCode)
    Parameter(3) = GenerateInputParameter("@Flag", adInteger, 4, 1)  'Update Records
    Parameter(4) = GenerateInputParameter("@BeforeDate", adVarWChar, 8, BeforeDate)
    Parameter(5) = GenerateInputParameter("@AfterDate", adVarWChar, 8, AfterDate)
    Parameter(6) = GenerateOutputParameter("@NumberOfRecords", adInteger, 4)
    
    NumberOfRecords = RunParametricStoredProcedure2String("Update_HavalehResid", Parameter)
    
    mdifrm.MousePointer = vbDefault
    frmDisMsg.lblMessage = " »Â —Ê“ “”«‰Ì ﬁÌ„  ÕÊ«·Â Â« Ê —”Ìœ Â« «‰Ã«„ ê—œÌœ "
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    UpdateHavaleResid = True
End Function

Public Function IntToHarf(ByRef No) As String
If Len(No) > 3 Then IntToHarf = "Œÿ«": Exit Function
Dim Mablagh, st1, st2, st3, StrNo    As String
StrNo = CStr(No)

If Len(StrNo) = 1 Then st1 = Right(StrNo, 1)
If Len(StrNo) = 2 Then st1 = Right(StrNo, 1): st2 = Mid(StrNo, 1, 1)
If Len(StrNo) = 3 Then st1 = Right(StrNo, 1): st2 = Mid(StrNo, 2, 1): st3 = Left(StrNo, 1)

Select Case st1
    Case "1": Mablagh = "Ìò"
    Case "2": Mablagh = "œÊ"
    Case "3": Mablagh = "”Â"
    Case "4": Mablagh = "çÂ«—"
    Case "5": Mablagh = "Å‰Ã"
    Case "6": Mablagh = "‘‘"
    Case "7": Mablagh = "Â› "
    Case "8": Mablagh = "Â‘ "
    Case "9": Mablagh = "‰Â"
End Select
If (st1 <> "0") And (st2 <> "") Then Mablagh = " Ê " + Mablagh

Select Case st2
    Case 1:
        Select Case st1
              Case "0":  Mablagh = "œÂ"
              Case "1":  Mablagh = "Ì«“œÂ"
              Case "2":  Mablagh = "œÊ«“œÂ"
              Case "3":  Mablagh = "”Ì“œÂ"
              Case "4":  Mablagh = "çÂ«—œÂ"
              Case "5":  Mablagh = "Å«‰“œÂ"
              Case "6":  Mablagh = "‘«‰“œÂ"
              Case "7":  Mablagh = "Â›œÂ"
              Case "8":  Mablagh = "ÂÃœÂ"
              Case "9":  Mablagh = "‰Ê“œÂ"
        End Select
    Case 2: Mablagh = "»Ì” " + Mablagh
    Case 3: Mablagh = "”Ì" + Mablagh
    Case 4: Mablagh = "çÂ·" + Mablagh
    Case 5: Mablagh = "Å‰Ã«Â" + Mablagh
    Case 6: Mablagh = "‘’ " + Mablagh
    Case 7: Mablagh = "Â› «œ" + Mablagh
    Case 8: Mablagh = "Â‘ «œ" + Mablagh
    Case 9: Mablagh = "‰Êœ" + Mablagh
End Select

  If (Mablagh <> "") And (st3 <> "") And (st2 <> "0") Then Mablagh = " Ê " + Mablagh

Select Case st3
    Case 1: Mablagh = "Ìò’œ" + Mablagh
    Case 2: Mablagh = "œÊÌ” " + Mablagh
    Case 3: Mablagh = "”Ì’œ" + Mablagh
    Case 4: Mablagh = "çÂ«—’œ" + Mablagh
    Case 5: Mablagh = "Å«‰’œ" + Mablagh
    Case 6: Mablagh = "‘‘’œ" + Mablagh
    Case 7: Mablagh = "Â› ’œ" + Mablagh
    Case 8: Mablagh = "Â‘ ’œ" + Mablagh
    Case 9: Mablagh = "‰Â’œ" + Mablagh
End Select

IntToHarf = Mablagh
End Function

Public Function NumberToHarf(No As Double) As String
    Dim Mablagh1, Mablagh2, Mablagh3, Mablagh4, st1, StrNo As String
    Dim i1, i2, i3, i4, Code As Integer
    StrNo = CStr(No)
    Mablagh1 = ""
    Mablagh2 = ""
    Mablagh3 = ""
    Mablagh4 = ""
    If Len(StrNo) > 12 Then NumberToHarf = "Œÿ«": Exit Function
    StrNo = Format(StrNo, "000000000000")
    i1 = Val(Mid(StrNo, 1, 3))
    i2 = Val(Mid(StrNo, 4, 3))
    i3 = Val(Mid(StrNo, 7, 3))
    i4 = Val(Mid(StrNo, 10, 3))
    If i1 > 0 Then Mablagh1 = IntToHarf(i1)
    If i2 > 0 Then Mablagh2 = IntToHarf(i2)
    If i3 > 0 Then Mablagh3 = IntToHarf(i3)
    If i4 > 0 Then Mablagh4 = IntToHarf(i4)
    If (Mablagh1 <> "") Then
        Mablagh1 = Mablagh1 + " „Ì·Ì«—œ"
        If (Mablagh2 <> "") Or (Mablagh3 <> "") Or (Mablagh4 <> "") Then Mablagh1 = Mablagh1 + " Ê "
    End If
    If (Mablagh2 <> "") Then

        Mablagh2 = Mablagh2 + " „Ì·ÌÊ‰"
        If (Mablagh3 <> "") Or (Mablagh4 <> "") Then Mablagh2 = Mablagh2 + " Ê "

    End If
    If (Mablagh3 <> "") Then

        Mablagh3 = Mablagh3 + " Â“«—"
        If (Mablagh4 <> "") Then Mablagh3 = Mablagh3 + " Ê "

    End If
    NumberToHarf = Mablagh1 + Mablagh2 + Mablagh3 + Mablagh4
    
End Function

Public Sub PosLogSave(LogDescription As String)
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim LogFile As String
    
    
    LogFile = App.Path & "\Logs\" & DateToNumber8(Right(clsDate.shamsi(Date), 8)) & ".Log"
    If Not filetemp.FolderExists(App.Path & "\Logs") Then
        filetemp.CreateFolder App.Path & "\Logs"
    End If
    If filetemp.FileExists(LogFile) Then
        Set tempstring = filetemp.OpenTextFile(LogFile, ForAppending, False, TristateFalse)
    Else
        filetemp.CreateTextFile LogFile
        Set tempstring = filetemp.OpenTextFile(LogFile, ForWriting, False, TristateFalse)
    End If
    
    tempstring.WriteBlankLines 2
    tempstring.Write (LogDescription + Chr(13))
    tempstring.WriteBlankLines 2
    tempstring.Close
   ' MsgBox " . –ŒÌ—Â ê—œÌœ " & LogFile & "›«Ì· „ ‰Ì œ—  "
End Sub


Public Function InsertPos_tfaccard(FactorNo As Long, Status As EnumFactorType, _
       PosNo As Long, NvcBatchNo As String, NvcTraceNo As String, CardAuthNumber As String, _
       CardNumber As String, TransTime As String) As Boolean
    
    On Error GoTo ErrorHandler
    ReDim pa(1 To 9) As Parameter
    pa(1) = GenerateInputParameter("@nf", adInteger, 4, FactorNo)
    pa(2) = GenerateInputParameter("@Status", adInteger, 4, Status)
    pa(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    pa(4) = GenerateInputParameter("@PosId", adInteger, 4, PosNo)
    pa(5) = GenerateInputParameter("@NvcBatchNo", adWChar, 20, NvcBatchNo)
    pa(6) = GenerateInputParameter("@NvcTraceNo", adWChar, 20, NvcTraceNo)
    pa(7) = GenerateInputParameter("@CardAuthNumber", adWChar, 20, CardAuthNumber)
    pa(8) = GenerateInputParameter("@CardNumber", adWChar, 20, Left(CardNumber, 6) & "******" & Right(CardNumber, 4))
    pa(9) = GenerateInputParameter("@TransTime", adWChar, 50, TransTime)
    RunParametricStoredProcedure "Update_tFacCard_ByTranDetials", pa
    
    InsertPos_tfaccard = True
    
    Exit Function

ErrorHandler:
    InsertPos_tfaccard = False
    
End Function

Public Sub CloseWindow(ByVal WindowName As String)
    
    Dim hWnd, retval As Long
    Dim WinTitle As String
    'WinTitle = "Recycle Bin" '<- Title of Window
    WinTitle = WindowName  '<- Title of Window
    hWnd = FindWindow(vbNullString, WinTitle)
    retval = PostMessage(hWnd, WM_CLOSE, 0&, 0&)

End Sub

Public Sub ChangewinTitle(ByVal WindowName As String, WindowNewName As String)
On Error Resume Next
    Dim hWnd, retval As Long
    Dim WinTitle As String
    WinTitle = WindowName  '<- Title of Window
    hWnd = FindWindow(vbNullString, WinTitle)
    Call SendMessageByString(hWnd, WM_SETTEXT, 0&, WindowNewName)
End Sub

Public Sub PresetScreenSaver()
    mdifrm.tmrScreenSaver.Enabled = False
    mdifrm.tmrScreenSaver.Interval = 60000
    timeInterval = clsInvoiceValue.ScreenSaverTime
    If timeInterval = 0 Then Exit Sub
    mdifrm.tmrScreenSaver.Enabled = True
End Sub


Public Sub SetColor()
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
    
    Dim IsFileExist As Boolean
    
    If UserSettingFile = "" Then End    'Only  For  Make Exe File
    Set tempstring = filetemp.OpenTextFile(UserSettingFile, ForReading, False, TristateFalse)
    
    Do While tempstring.AtEndOfLine = False
       Str = tempstring.ReadLine
       LenghStr = InStr(1, Str, "=", vbTextCompare)
       
       If InStr(1, Str, "Invoice_BackColorForm", vbTextCompare) Then
          Invoice_BackColorForm = Val(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "Invoice_BackColorBtn0", vbTextCompare) Then
          Invoice_BackColorBtn0 = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_BackColorBtn1", vbTextCompare) Then
          Invoice_BackColorBtn1 = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_BackColorBtn2", vbTextCompare) Then
          Invoice_BackColorBtn2 = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_BackColorBtn3", vbTextCompare) Then
          Invoice_BackColorBtn3 = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_BackColorBtn4", vbTextCompare) Then
          Invoice_BackColorBtn4 = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_BackColorBtn5", vbTextCompare) Then
          Invoice_BackColorBtn5 = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_BackColorBtn6", vbTextCompare) Then
          Invoice_BackColorBtn6 = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_BackColorBtn7", vbTextCompare) Then
          Invoice_BackColorBtn7 = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_BackColorFlexGrid", vbTextCompare) Then
          Invoice_BackColorFlexGrid = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_FontMenuName", vbTextCompare) Then
          Invoice_FontMenuName = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_FontMenuSize", vbTextCompare) Then
          Invoice_FontMenuSize = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_FontMenuBold", vbTextCompare) Then
          Invoice_FontMenuBold = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_FontFlexGridName", vbTextCompare) Then
          Invoice_FontFlexGridName = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_FontFlexGridSize", vbTextCompare) Then
          Invoice_FontFlexGridSize = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_FontFlexGridBold", vbTextCompare) Then
          Invoice_FontFlexGridBold = Mid(Str, LenghStr + 1)
       
       
       ElseIf InStr(1, Str, "Invoice_FontDifferencesName", vbTextCompare) Then
          Invoice_FontDifferencesName = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_FontDifferencesSize", vbTextCompare) Then
          Invoice_FontDifferencesSize = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Invoice_FontDifferencesBold", vbTextCompare) Then
          Invoice_FontDifferencesBold = Mid(Str, LenghStr + 1)
       
       End If
    Loop
    tempstring.Close
    
    If Invoice_BackColorBtn0 = 0 Then Invoice_BackColorBtn0 = 12640511
    If Invoice_BackColorBtn1 = 0 Then Invoice_BackColorBtn1 = 12640511
    If Invoice_BackColorBtn2 = 0 Then Invoice_BackColorBtn2 = 12640511
    If Invoice_BackColorBtn3 = 0 Then Invoice_BackColorBtn3 = 12640511

End Sub

Public Function GetFarsiStringFromArabic(instring As String) As String
    On Error GoTo ErrHandler

    Dim rctmp As New ADODB.Recordset
      
         ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@nvcMainString", adVarWChar, 4000, instring)
            Set rctmp = RunParametricStoredProcedure2Rec("Get_ArabicToFarsiStringSp", Parameter)
 
        Do While Not (rctmp.EOF)
            GetFarsiStringFromArabic = rctmp!Result
            rctmp.MoveNext
        Loop
        If rctmp.State = adStateOpen Then rctmp.Close
       
    Exit Function
ErrHandler:
    modgl.LogSaveNew "frmInvoice => ", err.Description, err.Number, err.Source, "GetFarsiStringFromArabic"
    ShowErrorMessage
End Function


Public Function GetFarsiHisenseCustomerDisplay(txt As String) As String
Dim fnt As Variant
Dim Inp(30), outp(30)
fnt = Array(Array(1570, 1, 71), Array(1575, 2, 72), Array(1576, 4, 74), Array(1662, 4, 38), _
            Array(1578, 4, 78), Array(1579, 4, 82), Array(1580, 4, 86), Array(1670, 4, 42), _
            Array(1581, 4, 90), Array(1582, 4, 94), Array(1583, 2, 98), Array(1584, 2, 100), _
            Array(1585, 2, 102), Array(1586, 2, 104), Array(1688, 2, 46), Array(1587, 4, 106), _
            Array(1588, 4, 110), Array(1589, 4, 114), Array(1590, 4, 118), Array(1591, 4, 122), _
            Array(1592, 4, 126), Array(1593, 4, 130), Array(1594, 4, 134), Array(1601, 4, 138), _
            Array(1602, 4, 142), Array(1705, 4, 146), Array(1711, 4, 48), Array(1604, 4, 150), _
            Array(1605, 4, 154), Array(1606, 4, 158), Array(1608, 2, 166), Array(1607, 4, 162), _
            Array(1740, 4, 168), Array(1574, 4, 67), Array(1548, 1, 54), Array(1567, 1, 66), _
            Array(48, 1, 55), Array(49, 1, 56), Array(50, 1, 57), Array(51, 1, 58), _
            Array(52, 1, 59), Array(53, 1, 60), Array(54, 1, 61), Array(55, 1, 62), _
            Array(56, 1, 63), Array(57, 1, 64))

            
Dim i As Long
Dim j As Integer
For i = 0 To Len(txt) - 1
If Mid(txt, i + 1, 1) <> " " Then
    If AscW(Mid(txt, i + 1, 1)) < 128 Then
    
        
       
        Inp(i) = 1000 + AscW(Mid(txt, i + 1, 1))
       
    Else
        For j = 0 To 35
            If AscW(Mid(txt, i + 1, 1)) = fnt(j)(0) Then Exit For
        Next
        Inp(i) = j
    End If
End If
If Mid(txt, i + 1, 1) = " " Then Inp(i) = 3600


Next
Dim farsicntr As Integer
farsicntr = 0
Dim Result As String
Result = ""
Dim ocntr As Integer
Dim icntr As Integer
Dim State As Integer

ocntr = 0
icntr = 0
Do While (ocntr < 20 And icntr < Len(txt))
   
    If Inp(icntr) > 1000 Then
        outp(ocntr) = Inp(icntr) - 1000 + 7
        State = 0
        farsicntr = 0
    Else
        If farsicntr = 0 Then
            outp(ocntr) = fnt(Inp(icntr))(2)
            State = 0
        ElseIf fnt(Inp(icntr - 1))(1) < 4 Then
            outp(ocntr) = fnt(Inp(icntr))(2)
            State = 0
        Else
            Select Case fnt(Inp(icntr))(1)
                Case 1
                    outp(ocntr) = fnt(Inp(icntr))(2)
                    State = 0
                Case 2
                    If Inp(icntr) = 1 And Inp(icntr - 1) = 27 Then  '''
                        If State = 0 Then                             '
                            outp(ocntr - 1) = 52                      '
                        ElseIf State = 3 Then                         ''' Check for la
                            outp(ocntr - 1) = 53                      '
                        End If                                        '
                        ocntr = ocntr - 1                           '''
                    Else
                        If State = 0 Then
                            outp(ocntr - 1) = outp(ocntr - 1) + 1
                        ElseIf State = 3 Then
                            outp(ocntr - 1) = outp(ocntr - 1) - 1
                        End If
                        outp(ocntr) = fnt(Inp(icntr))(2) + 1
                        State = 0   'State is not important.
                    End If
                Case 4
                    If State = 0 Then
                        outp(ocntr - 1) = outp(ocntr - 1) + 1
                    ElseIf State = 3 Then
                        outp(ocntr - 1) = outp(ocntr - 1) - 1
                    End If
                    outp(ocntr) = fnt(Inp(icntr))(2) + 3
                    State = 3
            End Select
        End If
        farsicntr = farsicntr + 1
    End If
    If outp(ocntr) = 32 Then
        ocntr = ocntr - 1
    Else
       'Cells(4, 2 + ocntr) = outp(ocntr)
        'If farsicntr > 1 Then Cells(4, 2 + ocntr - 1) = outp(ocntr - 1)
    End If

    ocntr = ocntr + 1
    icntr = icntr + 1

Loop
Dim Position As Integer
Result = ""
        For Position = LBound(outp) To UBound(outp)
            

            If outp(Position) <> 2607 And outp(Position) <> "" Then
'                If outp(Position) = "39" Then
'                Result = Result & Chr$(Val("&H" & CStr(Hex$("20")))) 'outp (Position)
'                Else
                Result = Result & Chr(Val("&H" & CStr(Hex$(outp(Position))))) 'outp (Position)
                End If
           ' End If
        Next Position

    
GetFarsiHisenseCustomerDisplay = Result

End Function

Public Sub AddStationtoDB(ByVal i As Long, StationType As Long)
    Dim cmd As New ADODB.command
    Dim rctmp As New ADODB.Recordset
    Dim MaxStationId As Long
    Dim j, k As Long
    If StationType = 2 Then k = clsArya.MaxStationNo Else k = clsArya.MaxPocketPcNo
    rctmp.Open "Select Isnull(max(StationId) , 0) as MaxStationId from tstations where Branch =  " & CurrentBranch, PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        MaxStationId = rctmp!MaxStationId
        With cmd
            .ActiveConnection = PosConnection
            For j = i To k
                 MaxStationId = MaxStationId + 1
                .CommandType = adCmdText
                .CommandText = "INSERT INTO dbo.tStations ( StationID,Description,IsActive,StationType,Branch)" & _
                               "VALUES ( " & MaxStationId & "  , 'Station" & MaxStationId & "',1 ," & StationType & "," & CurrentBranch & ")"
                .Execute
                .Cancel
            Next
            If StationType = 2 Then
                ShowDisMessage " ⁄œ«œ  " & k - i + 1 & "  «Ì” ê«Â ÃœÌœ œ—œÌ «»Ì”  ⁄—Ì› ê—œÌœ ", 1500
            Else
                ShowDisMessage " ⁄œ«œ  " & k - i + 1 & "  «Ì” ê«Â ÃœÌœ Å«ﬂ  ÅÌ ”Ì œ—œÌ «»Ì”  ⁄—Ì› ê—œÌœ ", 1500
            End If
        End With
    End If
    If rctmp.State = 1 Then rctmp.Close
    Set cmd = Nothing
End Sub

Public Sub AddServertoDB()
    Dim cmd As New ADODB.command
    Dim rctmp As New ADODB.Recordset
    Dim MaxStationId As Long
    rctmp.Open "Select Isnull(max(StationId) , 0) as MaxStationId from tstations where Branch =  " & CurrentBranch, PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        MaxStationId = rctmp!MaxStationId
        With cmd
            .ActiveConnection = PosConnection
             MaxStationId = MaxStationId + 1
            .CommandType = adCmdText
            .CommandText = "INSERT INTO dbo.tStations ( StationID,Description,IsActive,StationType,Branch)" & _
                           "VALUES ( " & MaxStationId & "  , 'Station" & MaxStationId & "',1 ,3," & CurrentBranch & ")"
            .Execute
            .Cancel
            ShowDisMessage "  ”—Ê— ÃœÌœ œ—œÌ «»Ì”  ⁄—Ì› ê—œÌœ ", 1500
        End With
    End If
    If rctmp.State = 1 Then rctmp.Close
    Set cmd = Nothing
End Sub


' Set a form always on the top.
'
' the form can be specified as a Form or object
' or through its hWnd property
' If OnTop=False the always on the top mode is de-activated.
Sub SetAlwaysOnTopMode(hWndOrForm As Variant, Optional ByVal onTop As Boolean = _
True)
Dim hWnd As Long
' get the hWnd of the form to be move on top
If VarType(hWndOrForm) = vbLong Then
hWnd = hWndOrForm
Else
hWnd = hWndOrForm.hWnd
End If
SetWindowPos hWnd, IIf(onTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, _
SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
End Sub

Public Sub OnTopMe(FormID As Object, onTop As Boolean)
     If onTop = True Then SetWindowPos FormID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
     If onTop = False Then SetWindowPos FormID.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


Public Function GetSystemPath()
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetSystemDirectory(strFolder, MAX_PATH)
If lngResult <> 0 Then
    GetSystemPath = Left(strFolder, InStr(strFolder, _
    Chr(0)) - 1)
Else
    GetSystemPath = ""
End If
End Function

Public Function GetSystemPath64()
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetSystemWow64Directory(strFolder, MAX_PATH)
If lngResult <> 0 Then
    GetSystemPath64 = Left(strFolder, InStr(strFolder, _
    Chr(0)) - 1)
Else
    GetSystemPath64 = ""
End If
End Function

Public Sub LoadForm(formName As String)
    Dim varForm As Form
    Dim frmact As Form
    
    For Each varForm In Forms
        If formName = varForm.Name Then
            Set frmact = varForm
            Exit For
        End If
    Next

    formloadFlag = False
    frmact.Left = Val(GetSetting(strMainKey, formName, "Left"))
    If Val(GetSetting(strMainKey, formName, "Height")) > 0 Then
        frmact.Height = Val(GetSetting(strMainKey, formName, "Height"))
    End If
    If Val(GetSetting(strMainKey, formName, "Width")) > 0 Then
        frmact.Width = Val(GetSetting(strMainKey, formName, "Width"))
    End If
    frmact.Top = Val(GetSetting(strMainKey, formName, "Top"))
    formloadFlag = True

End Sub

Public Sub UnLoadForm(formName As String)
    Dim varForm As Form
    Dim frmact As Form
    
    For Each varForm In Forms
        If formName = varForm.Name Then
            Set frmact = varForm
            Exit For
        End If
    Next
'    SaveSetting strMainKey, FormName, "Left", frmAct.Left
'    SaveSetting strMainKey, FormName, "Top", frmAct.Top
End Sub

Public Sub CenterCenterOffset(ByRef MyForm As Form)
    MyForm.Left = (Screen.Width - MyForm.Width - frmGroupMenu.Width) / 3
    MyForm.Top = (Screen.Height - MyForm.Height) / 3
End Sub

Sub Sendkey(Text$, Optional wait As Boolean = False)
    Dim WshShell As Object

     'wrapper for Sendkeys which does not crash in the IDE under Windows Vista
     Set WshShell = CreateObject("WScript.Shell")
     WshShell.SendKeys Text, wait
     Set WshShell = Nothing

 End Sub

Public Function CheckDate6Digit(strDate As Variant) As Boolean
    Dim y As String
    Dim m As String
    Dim D As String

    On Error GoTo ErrorHandle

    y = Mid(strDate, 1, 2)
    m = Mid(strDate, 4, 2)
    D = Right(strDate, 2)

    CheckDate6Digit = False
    If Len(strDate) < 8 Then
        Exit Function
    End If

    If Len(y) < 2 Or Len(m) < 2 Or Len(D) < 2 Then Exit Function

    Select Case Val(m)
        Case 1 To 6
            If D > 31 Or D < 1 Then
                MsgBox " «—ÌŒ —« œ—”  Ê«—œ ﬂ‰Ìœ", vbOKOnly, "Œÿ«"
                Exit Function
            End If
        Case 7 To 12
            If D > 30 Or D < 1 Then
                MsgBox " «—ÌŒ —« œ—”  Ê«—œ ﬂ‰Ìœ", vbOKOnly, "Œÿ«"
                Exit Function
            End If
        Case Else
            MsgBox " «—ÌŒ —« œ—”  Ê«—œ ﬂ‰Ìœ", vbOKOnly, "Œÿ«"
            Exit Function
    End Select
    CheckDate6Digit = True
    Exit Function

ErrorHandle:
    CheckDate6Digit = False
End Function

