VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicInvoiceValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_ColRow As Boolean
'Private m_ColAmount As Boolean
'Private m_ColGoodName As Boolean
Private m_ColFee As Boolean
Private m_ColTotal As Boolean
Private m_ColGoodCode As Boolean
'Private m_ColGoodWeight As Boolean
Private m_ColUnitGood As Boolean
'Private m_ColServe As Boolean
'Private m_ColDifferencesCodes As Boolean
Private m_ColChanges As Boolean
Private m_ColSeller As Boolean
Private m_ColDiscount As Boolean
Private m_ColRate As Boolean
'Private m_ColChair As Boolean
Private m_ColStore As Boolean
'Private m_ColNoOfUnit As Boolean
Private m_ColMojodi As Boolean
Private m_ColDuty As Boolean
Private m_ColTax As Boolean
Private m_ShowPictureGood As Boolean
Private m_ShowGoodTime As Long
Private m_ShowInvoiceMenu As Boolean
Private m_GoodMenuView As Boolean
Private m_GoodMenuFileName As String
Private m_ScreenSaverTime As Long
Private m_LanguageIcon As Boolean
Private m_KeyboardIcon As Boolean
Private m_ColorIcon As Boolean
Private m_TelephoneIcon As Boolean
Private m_ShowLogo As Boolean
Private m_PrintLable As Boolean

Public Property Get ColRow() As Boolean
    ColRow = m_ColRow
End Property
Public Property Let ColRow(ByVal vData As Boolean)
    m_ColRow = vData
End Property
Public Property Get ColFee() As Boolean
    ColFee = m_ColFee
End Property
Public Property Let ColFee(ByVal vData As Boolean)
    m_ColFee = vData
End Property
Public Property Get ColTotal() As Boolean
    ColTotal = m_ColTotal
End Property
Public Property Let ColTotal(ByVal vData As Boolean)
    m_ColTotal = vData
End Property
Public Property Get ColGoodCode() As Boolean
    ColGoodCode = m_ColGoodCode
End Property
Public Property Let ColGoodCode(ByVal vData As Boolean)
    m_ColGoodCode = vData
End Property
Public Property Get ColUnitGood() As Boolean
    ColUnitGood = m_ColUnitGood
End Property
Public Property Let ColUnitGood(ByVal vData As Boolean)
    m_ColUnitGood = vData
End Property
Public Property Get ColChanges() As Boolean
    ColChanges = m_ColChanges
End Property
Public Property Let ColChanges(ByVal vData As Boolean)
    m_ColChanges = vData
End Property
Public Property Get ColSeller() As Boolean
    ColSeller = m_ColSeller
End Property
Public Property Let ColSeller(ByVal vData As Boolean)
    m_ColSeller = vData
End Property
Public Property Get ColDiscount() As Boolean
    ColDiscount = m_ColDiscount
End Property
Public Property Let ColDiscount(ByVal vData As Boolean)
    m_ColDiscount = vData
End Property
Public Property Get ColRate() As Boolean
    ColRate = m_ColRate
End Property
Public Property Let ColRate(ByVal vData As Boolean)
    m_ColRate = vData
End Property
Public Property Get ColStore() As Boolean
    ColStore = m_ColStore
End Property
Public Property Let ColStore(ByVal vData As Boolean)
    m_ColStore = vData
End Property
Public Property Get ColMojodi() As Boolean
    ColMojodi = m_ColMojodi
End Property
Public Property Let ColMojodi(ByVal vData As Boolean)
    m_ColMojodi = vData
End Property
Public Property Get ColDuty() As Boolean
    ColDuty = m_ColDuty
End Property
Public Property Let ColDuty(ByVal vData As Boolean)
    m_ColDuty = vData
End Property
Public Property Get ColTax() As Boolean
    ColTax = m_ColTax
End Property
Public Property Let ColTax(ByVal vData As Boolean)
    m_ColTax = vData
End Property
Public Property Get ShowPictureGood() As Boolean
    ShowPictureGood = m_ShowPictureGood
End Property
Public Property Let ShowPictureGood(ByVal vData As Boolean)
    m_ShowPictureGood = vData
End Property
Public Property Get ShowGoodTime() As Long
    ShowGoodTime = m_ShowGoodTime
End Property
Public Property Let ShowGoodTime(ByVal vData As Long)
    m_ShowGoodTime = vData
End Property
Public Property Get ShowInvoiceMenu() As Boolean
    ShowInvoiceMenu = m_ShowInvoiceMenu
End Property
Public Property Let ShowInvoiceMenu(ByVal vData As Boolean)
    m_ShowInvoiceMenu = vData
End Property
Public Property Get GoodMenuView() As Boolean
    GoodMenuView = m_GoodMenuView
End Property
Public Property Let GoodMenuView(ByVal vData As Boolean)
    m_GoodMenuView = vData
End Property
Public Property Get GoodMenuFileName() As String
    GoodMenuFileName = m_GoodMenuFileName
End Property
Public Property Let GoodMenuFileName(ByVal vData As String)
    m_GoodMenuFileName = vData
End Property
Public Property Get ScreenSaverTime() As Long
    ScreenSaverTime = m_ScreenSaverTime
End Property
Public Property Let ScreenSaverTime(ByVal vData As Long)
    m_ScreenSaverTime = vData
End Property
Public Property Get LanguageIcon() As Long
    LanguageIcon = m_LanguageIcon
End Property
Public Property Let LanguageIcon(ByVal vData As Long)
    m_LanguageIcon = vData
End Property
Public Property Get KeyboardIcon() As Long
    KeyboardIcon = m_KeyboardIcon
End Property
Public Property Let KeyboardIcon(ByVal vData As Long)
    m_KeyboardIcon = vData
End Property
Public Property Get ColorIcon() As Long
    ColorIcon = m_ColorIcon
End Property
Public Property Let ColorIcon(ByVal vData As Long)
    m_ColorIcon = vData
End Property
Public Property Get TelephoneIcon() As Long
    TelephoneIcon = m_TelephoneIcon
End Property
Public Property Let TelephoneIcon(ByVal vData As Long)
    m_TelephoneIcon = vData
End Property
Public Property Get ShowLogo() As Long
    ShowLogo = m_ShowLogo
End Property
Public Property Let ShowLogo(ByVal vData As Long)
    m_ShowLogo = vData
End Property
Public Property Get PrintLable() As Boolean
    PrintLable = m_PrintLable
End Property
Public Property Let PrintLable(ByVal vData As Boolean)
    m_PrintLable = vData
End Property

Private Sub Class_Initialize()
    
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
    
    Dim IsFileExist As Boolean
    
    If InvoiceSettingFile = "" Then End    'Only  For  Make Exe File
    Set tempstring = filetemp.OpenTextFile(InvoiceSettingFile, ForReading, False, TristateFalse)
    
    Do While tempstring.AtEndOfLine = False
       Str = tempstring.ReadLine
       LenghStr = InStr(1, Str, "=", vbTextCompare)
       
       If InStr(1, Str, "ColRow", vbTextCompare) Then
          m_ColRow = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ColChanges", vbTextCompare) Then
          m_ColChanges = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ColDiscount", vbTextCompare) Then
          m_ColDiscount = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ColFee", vbTextCompare) Then
          m_ColFee = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ColGoodCode", vbTextCompare) Then
          m_ColGoodCode = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ColMojodi", vbTextCompare) Then
          m_ColMojodi = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ColRate", vbTextCompare) Then
          m_ColRate = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ColSeller", vbTextCompare) Then
          m_ColSeller = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ColStore", vbTextCompare) Then
          m_ColStore = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ColTotal", vbTextCompare) Then
          m_ColTotal = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ColUnitGood", vbTextCompare) Then
          m_ColUnitGood = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ColDuty", vbTextCompare) Then
          m_ColDuty = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ColTax", vbTextCompare) Then
          m_ColTax = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ShowPictureGood", vbTextCompare) Then
          m_ShowPictureGood = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ShowGoodTime", vbTextCompare) Then
          m_ShowGoodTime = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "ShowInvoiceMenu", vbTextCompare) Then
          m_ShowInvoiceMenu = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "GoodMenuView", vbTextCompare) Then
          m_GoodMenuView = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "GoodMenuFileName", vbTextCompare) Then
          m_GoodMenuFileName = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "ScreenSaverTime", vbTextCompare) Then
          m_ScreenSaverTime = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "LanguageIcon", vbTextCompare) Then
          m_LanguageIcon = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "KeyboardIcon", vbTextCompare) Then
          m_KeyboardIcon = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "ColorIcon", vbTextCompare) Then
          m_ColorIcon = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "TelephoneIcon", vbTextCompare) Then
          m_TelephoneIcon = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "ShowLogo", vbTextCompare) Then
          m_ShowLogo = Trim(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "PrintLable", vbTextCompare) Then
          m_PrintLable = Trim(Mid(Str, LenghStr + 1))
       
       End If
    Loop
    tempstring.Close

End Sub





