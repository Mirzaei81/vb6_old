Attribute VB_Name = "SettingFiles"

Public AryaSettingFile As String
Public StationSettingFile As String
Public UserSettingFile As String
Public InvoiceSettingFile As String
Public GoodMenuSettingFile As String
Public Invoice_FontMenuName As String
Public Invoice_FontMenuSize  As String
Public Invoice_FontMenuBold  As String
Public Invoice_FontDifferencesName As String
Public Invoice_FontDifferencesSize  As String
Public Invoice_FontDifferencesBold  As String
Public AccountingSettingFile As String
Public Invoice_BackColorForm As Long
Public Invoice_BackColorBtn0 As Long
Public Invoice_BackColorBtn1 As Long
Public Invoice_BackColorBtn2 As Long
Public Invoice_BackColorBtn3 As Long
Public Invoice_BackColorBtn4 As Long
Public Invoice_BackColorBtn5 As Long
Public Invoice_BackColorBtn6 As Long
Public Invoice_BackColorBtn7 As Long
Public Invoice_BackColorFlexGrid As Long
Public Invoice_FontFlexGridName As String
Public Invoice_FontFlexGridSize  As String
Public Invoice_FontFlexGridBold  As String

Public Function SetUserSettingFile(mvarColor As Long, index As Integer)

    Dim f As New FileSystemObject
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
     
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
       
       ElseIf InStr(1, Str, "Purchase_BackColorForm", vbTextCompare) Then
          Purchase_BackColorForm = Val(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "Purchase_BackColorBtn", vbTextCompare) Then
          Purchase_BackColorBtn = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Purchase_BackColorFlexGrid", vbTextCompare) Then
          Purchase_BackColorFlexGrid = Mid(Str, LenghStr + 1)
      
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
    
    Set tempstring = filetemp.OpenTextFile(UserSettingFile, ForWriting, False, TristateFalse)
    
    If Invoice_BackColorBtn0 = 0 Then Invoice_BackColorBtn0 = 12640511
    If Invoice_BackColorBtn1 = 0 Then Invoice_BackColorBtn1 = 12640511
    If Invoice_BackColorBtn2 = 0 Then Invoice_BackColorBtn2 = 12640511
    If Invoice_BackColorBtn3 = 0 Then Invoice_BackColorBtn3 = 12640511
    If Invoice_BackColorBtn4 = 0 Then Invoice_BackColorBtn4 = 12640511
    If Invoice_BackColorBtn5 = 0 Then Invoice_BackColorBtn5 = 12640511
    If Invoice_BackColorBtn6 = 0 Then Invoice_BackColorBtn6 = 12640511
    If Invoice_BackColorBtn7 = 0 Then Invoice_BackColorBtn7 = 12640511
    
    
    If index = 1 Then
        If VarActForm = "frmInvoice" Then
            Str = "Invoice_BackColorForm =" & mvarColor
            tempstring.WriteLine (Str)
            Str = "Purchase_BackColorForm =" & Purchase_BackColorForm
            tempstring.WriteLine (Str)
        ElseIf VarActForm = "frmPurchase" Then
            Str = "Purchase_BackColorForm =" & mvarColor
            tempstring.WriteLine (Str)
            Str = "Invoice_BackColorForm =" & Invoice_BackColorForm
            tempstring.WriteLine (Str)
        End If
        Str = "Invoice_BackColorBtn0 =" & Invoice_BackColorBtn0
        tempstring.WriteLine (Str)
        Str = "Invoice_BackColorBtn1 =" & Invoice_BackColorBtn1
        tempstring.WriteLine (Str)
        Str = "Invoice_BackColorBtn2 =" & Invoice_BackColorBtn2
        tempstring.WriteLine (Str)
        Str = "Invoice_BackColorBtn3 =" & Invoice_BackColorBtn3
        tempstring.WriteLine (Str)
        Str = "Invoice_BackColorBtn4 =" & Invoice_BackColorBtn4
        tempstring.WriteLine (Str)
        Str = "Invoice_BackColorBtn5 =" & Invoice_BackColorBtn5
        tempstring.WriteLine (Str)
        Str = "Invoice_BackColorBtn6 =" & Invoice_BackColorBtn6
        tempstring.WriteLine (Str)
        Str = "Invoice_BackColorBtn7 =" & Invoice_BackColorBtn7
        tempstring.WriteLine (Str)
        Str = "Invoice_BackColorFlexGrid =" & Invoice_BackColorFlexGrid
        tempstring.WriteLine (Str)
        Str = "Purchase_BackColorBtn =" & Purchase_BackColorBtn
        tempstring.WriteLine (Str)
        Str = "Purchase_BackColorFlexGrid =" & Purchase_BackColorFlexGrid
        tempstring.WriteLine (Str)
    ElseIf index = 20 Or index = 21 Or index = 22 Or index = 23 Or index = 24 Or index = 25 Or index = 26 Or index = 27 Then
        If index = 20 Then
            Str = "Invoice_BackColorBtn0 =" & mvarColor
            tempstring.WriteLine (Str)
        Else
            Str = "Invoice_BackColorBtn0 =" & Invoice_BackColorBtn0
            tempstring.WriteLine (Str)
        End If
        If index = 21 Then
            Str = "Invoice_BackColorBtn1 =" & mvarColor
            tempstring.WriteLine (Str)
        Else
            Str = "Invoice_BackColorBtn1 =" & Invoice_BackColorBtn1
            tempstring.WriteLine (Str)
        End If
        If index = 22 Then
            Str = "Invoice_BackColorBtn2 =" & mvarColor
            tempstring.WriteLine (Str)
        Else
            Str = "Invoice_BackColorBtn2 =" & Invoice_BackColorBtn2
            tempstring.WriteLine (Str)
        End If
        If index = 23 Then
            Str = "Invoice_BackColorBtn3 =" & mvarColor
            tempstring.WriteLine (Str)
        Else
            Str = "Invoice_BackColorBtn3 =" & Invoice_BackColorBtn3
            tempstring.WriteLine (Str)
        End If
        If index = 24 Then
            Str = "Invoice_BackColorBtn4 =" & mvarColor
            tempstring.WriteLine (Str)
        Else
            Str = "Invoice_BackColorBtn4 =" & Invoice_BackColorBtn4
            tempstring.WriteLine (Str)
        End If
        If index = 25 Then
            Str = "Invoice_BackColorBtn5 =" & mvarColor
            tempstring.WriteLine (Str)
        Else
            Str = "Invoice_BackColorBtn5 =" & Invoice_BackColorBtn5
            tempstring.WriteLine (Str)
        End If
        If index = 26 Then
            Str = "Invoice_BackColorBtn6 =" & mvarColor
            tempstring.WriteLine (Str)
        Else
            Str = "Invoice_BackColorBtn6 =" & Invoice_BackColorBtn6
            tempstring.WriteLine (Str)
        End If
        If index = 27 Then
            Str = "Invoice_BackColorBtn7 =" & mvarColor
            tempstring.WriteLine (Str)
        Else
            Str = "Invoice_BackColorBtn7 =" & Invoice_BackColorBtn7
            tempstring.WriteLine (Str)
        End If
        Str = "Invoice_BackColorBtn =" & Invoice_BackColorBtn
        tempstring.WriteLine (Str)
        Str = "Purchase_BackColorBtn =" & Purchase_BackColorBtn
        tempstring.WriteLine (Str)
        Str = "Invoice_BackColorForm =" & Invoice_BackColorForm
        tempstring.WriteLine (Str)
        Str = "Invoice_BackColorFlexGrid =" & Invoice_BackColorFlexGrid
        tempstring.WriteLine (Str)
        Str = "Purchase_BackColorForm =" & Purchase_BackColorForm
        tempstring.WriteLine (Str)
        Str = "Purchase_BackColorFlexGrid =" & Purchase_BackColorFlexGrid
        tempstring.WriteLine (Str)
    ElseIf index = 3 Then
        If VarActForm = "frmInvoice" Then
            Str = "Invoice_BackColorFlexGrid =" & mvarColor
            tempstring.WriteLine (Str)
            Str = "Purchase_BackColorFlexGrid =" & Purchase_BackColorFlexGrid
            tempstring.WriteLine (Str)
        ElseIf VarActForm = "frmPurchase" Then
            Str = "Purchase_BackColorFlexGrid =" & mvarColor
            tempstring.WriteLine (Str)
            Str = "Invoice_BackColorFlexGrid =" & Invoice_BackColorFlexGrid
            tempstring.WriteLine (Str)
        End If
        Str = "Invoice_BackColorForm =" & Invoice_BackColorForm
        tempstring.WriteLine (Str)
        Str = "Invoice_BackColorBtn =" & Invoice_BackColorBtn
        tempstring.WriteLine (Str)
        Str = "Purchase_BackColorForm =" & Purchase_BackColorForm
        tempstring.WriteLine (Str)
        Str = "Purchase_BackColorBtn =" & Purchase_BackColorBtn
        tempstring.WriteLine (Str)
    End If
    Str = "Invoice_FontMenuName =" & Invoice_FontMenuName
    tempstring.WriteLine (Str)
    Str = "Invoice_FontMenuSize =" & Invoice_FontMenuSize
    tempstring.WriteLine (Str)
    Str = "Invoice_FontMenuBold =" & Invoice_FontMenuBold
    tempstring.WriteLine (Str)
    Str = "Invoice_FontFlexGridName =" & Invoice_FontFlexGridName
    tempstring.WriteLine (Str)
    Str = "Invoice_FontFlexGridSize =" & Invoice_FontFlexGridSize
    tempstring.WriteLine (Str)
    Str = "Invoice_FontFlexGridBold =" & Invoice_FontFlexGridBold
    tempstring.WriteLine (Str)
    Str = "Invoice_FontDifferencesName =" & Invoice_FontDifferencesName
    tempstring.WriteLine (Str)
    Str = "Invoice_FontDifferencesSize =" & Invoice_FontDifferencesSize
    tempstring.WriteLine (Str)
    Str = "Invoice_FontDifferencesBold =" & Invoice_FontDifferencesBold
    tempstring.WriteLine (Str)
   

    tempstring.Close

End Function

Public Function SetStationSettingFile()

    Dim f As New FileSystemObject
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
     
    Set tempstring = filetemp.OpenTextFile(StationSettingFile, ForWriting, False, TristateFalse)
    
    Str = "PartitionId =" & clsStation.PartitionID
    tempstring.WriteLine (Str)
    
    Str = "ServePlaceDefault =" & clsStation.ServePlaceDefault
    tempstring.WriteLine (Str)
    
    Str = "PurchaseInventoryDefault =" & clsStation.PurchaseInventoryDefault
    tempstring.WriteLine (Str)
    
    Str = "WinAscii =" & clsStation.WinAscii
    tempstring.WriteLine (Str)
    
    Str = "MaxAutoDiscount =" & clsStation.MaxAutoDiscount
    tempstring.WriteLine (Str)
    
    Str = "Language =" & clsStation.Language
    tempstring.WriteLine (Str)
    
    Str = "DefaultCustSearch =" & clsStation.DefaultCustSearch
    tempstring.WriteLine (Str)
    
    Str = "DeliveryNoView =" & clsStation.DeliveryNoView
    tempstring.WriteLine (Str)
    
    Str = "AutoDrawerOpen =" & clsStation.AutoDrawerOpen
    tempstring.WriteLine (Str)
    
    Str = "ChangeGoodPrint =" & clsStation.ChangeGoodPrint
    tempstring.WriteLine (Str)
       
    Str = "AlphabeticGoods =" & clsStation.AlphabeticGoods
    tempstring.WriteLine (Str)
    
    Str = "TableControl =" & clsStation.TableControl
    tempstring.WriteLine (Str)
    
    Str = "RoundTwoNumber =" & clsStation.RoundTwoNumber
    tempstring.WriteLine (Str)
    
    Str = "DeliveryBarcodeDefault =" & clsStation.DeliveryBarcodeDefault
    tempstring.WriteLine (Str)
    
    Str = "TableBarcodeDefault =" & clsStation.TableBarcodeDefault
    tempstring.WriteLine (Str)
    
    Str = "ReprintDefault =" & clsStation.ReprintDefault
    tempstring.WriteLine (Str)
    
    Str = "KeyboardType =" & clsStation.KeyboardType
    tempstring.WriteLine (Str)
    
    Str = "SearchType =" & clsStation.SearchType
    tempstring.WriteLine (Str)
    
    Str = "PriceType =" & clsStation.PriceType
    tempstring.WriteLine (Str)
    
    Str = "MaxPrices =" & clsStation.MaxPrices
    tempstring.WriteLine (Str)
    
    Str = "DeletedGood =" & clsStation.DeletedGood
    tempstring.WriteLine (Str)
    
    Str = "SearchFichDefault =" & clsStation.SearchFichDefault
    tempstring.WriteLine (Str)
    
    Str = "CustomerOrderDefault =" & clsStation.CustomerOrderDefault
    tempstring.WriteLine (Str)
    
    Str = "CustomerServeplace =" & clsStation.CustomerServeplace
    tempstring.WriteLine (Str)
    
    Str = "CustomerSearchDefault =" & clsStation.CustomerSearchDefault
    tempstring.WriteLine (Str)
    
    Str = "CreditCalculate =" & clsStation.CreditCalculate
    tempstring.WriteLine (Str)
    
    Str = "GoodSearchDefault =" & clsStation.GoodSearchDefault
    tempstring.WriteLine (Str)
    
    Str = "DiscountDefault =" & clsStation.DiscountDefault
    tempstring.WriteLine (Str)
    
    Str = "SrarchInputDelayKeyboard =" & clsStation.SrarchInputDelayKeyboard
    tempstring.WriteLine (Str)
    
    Str = "MaxRecordCount =" & clsStation.MaxRecordCount
    tempstring.WriteLine (Str)
   
    Str = "FactorSortItems =" & clsStation.FactorSortItems
    tempstring.WriteLine (Str)
   
    Str = "MojodiControlDefault =" & clsStation.MojodiControlDefault
    tempstring.WriteLine (Str)
    
    Str = "RowMojodiControl =" & clsStation.RowMojodiControl
    tempstring.WriteLine (Str)
    
'    Str = "CommandView =" & clsStation.CommandView
'    tempstring.WriteLine (Str)
'
    Str = "EscapeInvoiceFactor =" & clsStation.EscapeInvoiceFactor
    tempstring.WriteLine (Str)
    
    Str = "StartUpFormDefault =" & clsStation.StartUpFormDefault
    tempstring.WriteLine (Str)
    
    Str = "Barcodelengh =" & clsStation.BarcodeLengh
    tempstring.WriteLine (Str)
    
    Str = "BarcodeChance =" & clsStation.BarcodeChance
    tempstring.WriteLine (Str)
    
    If clsStation.PriceChance = "" Then clsStation.PriceChance = "50000"
    Str = "PriceChance =" & clsStation.PriceChance
    tempstring.WriteLine (Str)
    
    Str = "RefreshFichNo =" & clsStation.RefreshFichNo
    tempstring.WriteLine (Str)
    
    Str = "DirectBascule =" & clsStation.DirectBascule
    tempstring.WriteLine (Str)
    
    Str = "StopOnEditFich =" & clsStation.StopOnEditFich
    tempstring.WriteLine (Str)
    
    Str = "MultiPrice =" & clsStation.MultiPrice
    tempstring.WriteLine (Str)
    
    Str = "UpdateBuyPrice =" & clsStation.UpdateBuyPrice
    tempstring.WriteLine (Str)
    
    Str = "UpdateSellPrice =" & clsStation.UpdateSellprice
    tempstring.WriteLine (Str)
    
    Str = "TrazooBarcode =" & clsStation.TrazooBarcode
    tempstring.WriteLine (Str)
    
    Str = "CodeFlag =" & clsStation.CodeFlag
    tempstring.WriteLine (Str)
    
    Str = "SoundAlarm =" & clsStation.SoundAlarm
    tempstring.WriteLine (Str)
    
    Str = "TypeBascule =" & clsStation.TypeBascule
    tempstring.WriteLine (Str)
    
    Str = "CashPayment =" & CBool(clsStation.CashPayment)
    tempstring.WriteLine (Str)
    
    Str = "InpersonSalonPayment =" & CBool(clsStation.InpersonSalonPayment)
    tempstring.WriteLine (Str)
    
    Str = "InpersonDeliveryPayment =" & CBool(clsStation.InpersonDeliveryPayment)
    tempstring.WriteLine (Str)
    
    Str = "InpersonOutPayment =" & CBool(clsStation.InpersonOutPayment)
    tempstring.WriteLine (Str)
    
    Str = "InpersonTablePayment =" & CBool(clsStation.InpersonTablePayment)
    tempstring.WriteLine (Str)
    
    Str = "InpersonSalonBalance =" & CBool(clsStation.InpersonSalonBalance)
    tempstring.WriteLine (Str)
    
    Str = "InpersonDeliveryBalance =" & CBool(clsStation.InpersonDeliveryBalance)
    tempstring.WriteLine (Str)
    
    Str = "InpersonOutBalance =" & CBool(clsStation.InpersonOutBalance)
    tempstring.WriteLine (Str)
    
    Str = "InpersonTableBalance =" & CBool(clsStation.InpersonTableBalance)
    tempstring.WriteLine (Str)
    
    Str = "ByPhoneSalonPayment =" & CBool(clsStation.ByPhoneSalonPayment)
    tempstring.WriteLine (Str)
    
    Str = "ByPhoneDeliveryPayment =" & CBool(clsStation.ByPhoneDeliveryPayment)
    tempstring.WriteLine (Str)
    
    Str = "ByPhoneTablePayment =" & CBool(clsStation.ByPhoneTablePayment)
    tempstring.WriteLine (Str)
    
    Str = "ByPhoneSalonBalance =" & CBool(clsStation.ByPhoneSalonBalance)
    tempstring.WriteLine (Str)
    
    Str = "ByPhoneDeliveryBalance =" & CBool(clsStation.ByPhoneDeliveryBalance)
    tempstring.WriteLine (Str)
    
    Str = "ByPhoneTableBalance =" & CBool(clsStation.ByPhoneTableBalance)
    tempstring.WriteLine (Str)
    
    Str = "ThreeSegmentSearch =" & CBool(clsStation.ThreeSegmentSearch)
    tempstring.WriteLine (Str)
    
    Str = "NumberOfUnitSale =" & CBool(clsStation.NumberOfUnitSale)
    tempstring.WriteLine (Str)
    
    Str = "MenuViewAfterGood =" & CBool(clsStation.MenuViewAfterGood)
    tempstring.WriteLine (Str)
    
    Str = "PayFactorView =" & CBool(clsStation.PayFactorView)
    tempstring.WriteLine (Str)
    
    Str = "FromStoreFee =" & clsStation.FromStoreFee
    tempstring.WriteLine (Str)
        
    Str = "BarcodeAutoEscape =" & CBool(clsStation.BarcodeAutoEscape)
    tempstring.WriteLine (Str)
    
    Str = "PrintAfterTasvieh =" & CBool(clsStation.PrintAfterTasvieh)
    tempstring.WriteLine (Str)
    
    Str = "AutoBarcode =" & CBool(clsStation.AutoBarcode)
    tempstring.WriteLine (Str)
    
    Str = "AutoCallerId =" & CBool(clsStation.AutoCallerId)
    tempstring.WriteLine (Str)
    
    Str = "CustomerAscii =" & CBool(clsStation.CustomerAscii)
    tempstring.WriteLine (Str)
    
    Str = "CustomerFarsi =" & CBool(clsStation.CustomerFarsi)
    tempstring.WriteLine (Str)
    
    Str = "CustomerOnlinePrice =" & CBool(clsStation.CustomerOnlinePrice)
    tempstring.WriteLine (Str)
    
    Str = "NoCurrentDay =" & CBool(clsStation.NoCurrentDay)
    tempstring.WriteLine (Str)
    
    Str = "SaleStartDefault =" & clsStation.SaleStartDefault
    tempstring.WriteLine (Str)
    
    Str = "ShowDigitNumber =" & clsStation.ShowDigitNumber
    tempstring.WriteLine (Str)
    
    Str = "CustomerRate =" & clsStation.CustomerRate
    tempstring.WriteLine (Str)
    
    Str = "GoodPercentage =" & CBool(clsStation.GoodPercentage)
    tempstring.WriteLine (Str)
    
    Str = "SearchOrderType =" & clsStation.SearchOrderType
    tempstring.WriteLine (Str)
    
    Str = "CallerIdSpace =" & clsStation.CallerIdSpace
    tempstring.WriteLine (Str)

'''    str = "Callwaiting =" & clsStation.Callwaiting
'''    tempstring.WriteLine (str)
    
    Str = "CountCustomerDailyBuy =" & clsStation.CountCustomerDailyBuy
    tempstring.WriteLine (Str)
   
    Str = "CountCustomerGood =" & clsStation.CountCustomerGood
    tempstring.WriteLine (Str)
    
    Str = "InvoiceRows =" & clsStation.InvoiceRows
    tempstring.WriteLine (Str)
   
    Str = "PurchaseRows =" & clsStation.PurchaseRows
    tempstring.WriteLine (Str)
   
    Str = "FinalCheck =" & CBool(clsStation.FinalCheck)
    tempstring.WriteLine (Str)
    
    Str = "AutoTip =" & CBool(clsStation.AutoTip)
    tempstring.WriteLine (Str)
    
    Str = "FichStatusBar =" & CBool(clsStation.FichStatusBar)
    tempstring.WriteLine (Str)
    
    Str = "CycleStockNoDefault =" & clsStation.CycleStockNoDefault
    tempstring.WriteLine (Str)
    
    Str = "TextIconViewH =" & CBool(clsStation.TextIconViewH)
    tempstring.WriteLine (Str)
     
    Str = "TextIconViewV =" & CBool(clsStation.TextIconViewv)
    tempstring.WriteLine (Str)
    
    Str = "ForceSeller =" & CBool(clsStation.ForceSeller)
    tempstring.WriteLine (Str)
    
    Str = "InvoiceStatusDefault =" & CBool(clsStation.InvoiceStatusDefault)
    tempstring.WriteLine (Str)
    
    Str = "NumberOfId =" & clsStation.NumberOfId
    tempstring.WriteLine (Str)
    
    Str = "DiscoveryPort =" & clsStation.DiscoveryPort
    tempstring.WriteLine (Str)
    
    Str = "ResponsePort =" & clsStation.ResponsePort
    tempstring.WriteLine (Str)
    
    Str = "CityCode =" & clsStation.CityCode
    tempstring.WriteLine (Str)
    
    Str = "AlmLogFile =" & CBool(clsStation.AlmLogFile)
    tempstring.WriteLine (Str)
    
    Str = "NetworkCallerId =" & CBool(clsStation.NetworkCallerId)
    tempstring.WriteLine (Str)
    
    Str = "TzWeightBarcodeId =" & clsStation.TzWeightBarcodeId
    tempstring.WriteLine (Str)
    
    Str = "CustomerFeeDataBase =" & clsStation.CustomerFeeDataBase
    tempstring.WriteLine (Str)
    
    Str = "GoodLevelAutoDiscount =" & clsStation.GoodLevelAutoDiscount
    tempstring.WriteLine (Str)
    
    Str = "StartNumberCartReader =" & clsStation.StartNumberCartReader
    tempstring.WriteLine (Str)
    
    Str = "NumberOfCardReader =" & clsStation.NumberOfCardReader
    tempstring.WriteLine (Str)
    
    Str = "ReadDirectWeightTz1 =" & clsStation.ReadDirectWeightTz1
    tempstring.WriteLine (Str)
    
    Str = "TaxView =" & clsStation.TaxView
    tempstring.WriteLine (Str)
    
    Str = "OutPrice =" & clsStation.OutPrice
    tempstring.WriteLine (Str)
    
    Str = "TzNumericBarcodeId =" & clsStation.TzNumericBarcodeId
    tempstring.WriteLine (Str)
    
    Str = "NumberOfUnitBuy =" & CBool(clsStation.NumberOfUnitBuy)
    tempstring.WriteLine (Str)
    
    Str = "EditCompatibleSamar1 =" & CBool(clsStation.EditCompatibleSamar1)
    tempstring.WriteLine (Str)
    
    Str = "UndoRedoCompatibleSamar1 =" & CBool(clsStation.UndoRedoCompatibleSamar1)
    tempstring.WriteLine (Str)
    
    Str = "EscNotExit =" & clsStation.EscNotExit
    tempstring.WriteLine (Str)
    
    Str = "ViewTempAddress =" & clsStation.ViewTempAddress
    tempstring.WriteLine (Str)
    
    Str = "CancelNotExit =" & clsStation.CancelNotExit
    tempstring.WriteLine (Str)
    
    Str = "ResiveFormerFichCurrentDay =" & CBool(clsStation.ResiveFormerFichCurrentDay)
    tempstring.WriteLine (Str)
    
    Str = "SelectInventory =" & CBool(clsStation.SelectInventory)
    tempstring.WriteLine (Str)
    
    Str = "AlphabetGoodSearch =" & clsStation.AlphabetGoodSearch
    tempstring.WriteLine (Str)
    
    Str = "DeficitLog =" & CBool(clsStation.DeficitLog)
    tempstring.WriteLine (Str)
    
    Str = "CallerId8Port =" & CBool(clsStation.CallerId8Port)
    tempstring.WriteLine (Str)
    
    Str = "ShiftRate =" & clsStation.ShiftRate
    tempstring.WriteLine (Str)
    
    Str = "SellerCaption =" & clsStation.SellerCaption
    tempstring.WriteLine (Str)
    
    Str = "SelectSeller =" & CBool(clsStation.SelectSeller)
    tempstring.WriteLine (Str)
    
    Str = "CountCustomerShiftBuy =" & clsStation.CountCustomerShiftBuy
    tempstring.WriteLine (Str)
    
    Str = "PrintAfterDeliver =" & CBool(clsStation.PrintAfterDeliver)
    tempstring.WriteLine (Str)
    
    Str = "FixRateChange =" & clsStation.FixRateChange
    tempstring.WriteLine (Str)
    
    Str = "AutoCashClose =" & clsStation.AutoCashClose
    tempstring.WriteLine (Str)
    
    Str = "PrintAfterDeliver =" & CBool(clsStation.PrintAfterDeliver)
    tempstring.WriteLine (Str)
    
    Str = "UpDateCarryFee =" & CBool(clsStation.UpDateCarryFee)
    tempstring.WriteLine (Str)
    
    Str = "ShowClock =" & CBool(clsStation.ShowClock)
    tempstring.WriteLine (Str)
    
    Str = "VoiceRecord =" & CBool(clsStation.VoiceRecord)
    tempstring.WriteLine (Str)
    
    Str = "TreeViewMenu =" & CBool(clsStation.TreeViewMenu)
    tempstring.WriteLine (Str)
    
    Str = "CallerIdAutoView =" & CBool(clsStation.CallerIdAutoView)
    tempstring.WriteLine (Str)
    
    Str = "CallerIdTest =" & CBool(clsStation.CallerIdTest)
    tempstring.WriteLine (Str)
    
    Str = "Pager =" & CBool(clsStation.Pager)
    tempstring.WriteLine (Str)
    
    Str = "AutoBackup =" & clsStation.AutoBackup
    tempstring.WriteLine (Str)
    
    Str = "PosPayment =" & CBool(clsStation.PosPayment)
    tempstring.WriteLine (Str)
    
    If clsStation.PassPhrase = "" Then clsStation.PassPhrase = "10000"
    Str = "PassPhrase =" & clsStation.PassPhrase
    tempstring.WriteLine (Str)
    
    If clsStation.PosApprovedText = "" Then clsStation.PosApprovedText = "<Approved>APPROVED</Approved>"
    Str = "PosApprovedText =" & clsStation.PosApprovedText
    tempstring.WriteLine (Str)
    
    Str = "PosModel =" & clsStation.PosModel
    tempstring.WriteLine (Str)
    
    Str = "ReportHeadername =" & clsStation.ReportHeadername
    tempstring.WriteLine (Str)
    
    Str = "FastCustSave =" & CBool(clsStation.FastCustSave)
    tempstring.WriteLine (Str)
    
    Str = "RepetitiveGood =" & clsStation.RepetitiveGood
    tempstring.WriteLine (Str)
    
    Str = "OtherPartition =" & clsStation.OtherPartition
    tempstring.WriteLine (Str)
    
    Str = "NotShowPrintNotice =" & clsStation.NotShowPrintNotice
    tempstring.WriteLine (Str)
    
    Str = "MultiInventory =" & clsStation.MultiInventory
    tempstring.WriteLine (Str)
    
    Str = "Frame_Printers =" & clsStation.Frame_Printers
    tempstring.WriteLine (Str)
    
    Str = "HasOptionPrice =" & clsStation.HasOptionPrice
    tempstring.WriteLine (Str)
    
    Str = "ShowOption =" & clsStation.ShowOption
    tempstring.WriteLine (Str)
    
    Str = "LableUsedGood =" & clsStation.LableUsedGood
    tempstring.WriteLine (Str)
    
    Str = "LoyaltyCustomers =" & clsStation.LoyaltyCustomers
    tempstring.WriteLine (Str)
    
    Str = "LoyaltyAllCustomers =" & clsStation.LoyaltyAllCustomers
    tempstring.WriteLine (Str)
    
    Str = "StartCharacter =" & clsStation.StartCharacter
    tempstring.WriteLine (Str)
    
    Str = "OneClickShow =" & clsStation.OneClickShow
    tempstring.WriteLine (Str)
    
    Str = "TouchScreen =" & clsStation.TouchScreen
    tempstring.WriteLine (Str)
    
    Str = "NoRowMenu =" & clsStation.NoRowMenu
    tempstring.WriteLine (Str)
    
    Str = "RfidReader =" & clsStation.RfidReader
    tempstring.WriteLine (Str)
    
    Str = "RfidInterval =" & clsStation.RfidInterval
    tempstring.WriteLine (Str)
    
    Str = "RfidLongBuzzer =" & clsStation.RfidLongBuzzer
    tempstring.WriteLine (Str)
    
    Str = "TelNetServerActive =" & clsStation.TelNetServerActive
    tempstring.WriteLine (Str)
    
    Str = "TelNetServerIP =" & clsStation.TelNetServerIP
    tempstring.WriteLine (Str)
    
    Str = "TelNetServerPort =" & clsStation.TelNetServerPort
    tempstring.WriteLine (Str)
    
    Str = "PrintAfterPayk =" & CBool(clsStation.PrintAfterPayk)
    tempstring.WriteLine (Str)
    
    Str = "PrintAfterOrder =" & CBool(clsStation.PrintAfterOrder)
    tempstring.WriteLine (Str)
    
    Str = "LabelPrint =" & clsStation.LabelPrint
    tempstring.WriteLine (Str)
    
    Str = "AryaSmsPanel =" & clsStation.AryaSmsPanel
    tempstring.WriteLine (Str)
    
    Str = "ForceTax =" & clsStation.ForceTax
    tempstring.WriteLine (Str)
    
    Str = "PersonIdCheck =" & clsStation.PersonIdCheck
    tempstring.WriteLine (Str)
    
    Str = "PersonIdRefreshTime =" & clsStation.PersonIdRefreshTime
    tempstring.WriteLine (Str)
    
    Str = "ListAutoLoad =" & clsStation.ListAutoLoad
    tempstring.WriteLine (Str)

    Str = "DeviceIP =" & clsStation.DeviceIP
    tempstring.WriteLine (Str)
    
    Str = "DeviceID =" & clsStation.DeviceID
    tempstring.WriteLine (Str)
    
    Str = "ListFont =" & clsStation.ListFont
    tempstring.WriteLine (Str)
    
    Str = "Device2IP =" & clsStation.Device2IP
    tempstring.WriteLine (Str)
    
    Str = "Device2Id =" & clsStation.Device2Id
    tempstring.WriteLine (Str)
    
    tempstring.Close

End Function

Public Function SetDefaultStationSettingFile()

    Dim f As New FileSystemObject
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
    
    f.CreateTextFile StationSettingFile
      
    Set tempstring = filetemp.OpenTextFile(StationSettingFile, ForWriting, False, TristateFalse)
    
    Str = "PartitionId =1"
    tempstring.WriteLine (Str)
    
    Str = "ServePlaceDefault =1"
    tempstring.WriteLine (Str)
    
    Str = "PurchaseInventoryDefault =1"
    tempstring.WriteLine (Str)
    
    Str = "WinAscii =1"
    tempstring.WriteLine (Str)
    
    Str = "Language =0"
    tempstring.WriteLine (Str)
    
    Str = "DefaultCustSearch =0"
    tempstring.WriteLine (Str)
    
    Str = "DeliveryBarcodeDefault =0"
    tempstring.WriteLine (Str)
    
    Str = "TableBarcodeDefault =0"
    tempstring.WriteLine (Str)
    
    Str = "ReprintDefault =0"
    tempstring.WriteLine (Str)
    
    Str = "MaxAutoDiscount =100"
    tempstring.WriteLine (Str)
    
    Str = "DeliveryNoView =False"
    tempstring.WriteLine (Str)
    
    Str = "AutoDrawerOpen =False"
    tempstring.WriteLine (Str)
    
    Str = "ChangeGoodPrint =False"
    tempstring.WriteLine (Str)
     
    Str = "AlphabeticGoods =False"
    tempstring.WriteLine (Str)
    
    Str = "TableControl =True"
    tempstring.WriteLine (Str)
    
    Str = "RoundTwoNumber =True"
    tempstring.WriteLine (Str)
    
    Str = "KeyboardType =0"
    tempstring.WriteLine (Str)
    
    Str = "SearchType =0"
    tempstring.WriteLine (Str)
    
    Str = "PriceType =1"
    tempstring.WriteLine (Str)
    
    Str = "MaxPrices =1"
    tempstring.WriteLine (Str)
    
    Str = "DeletedGood =True"
    tempstring.WriteLine (Str)
    
    Str = "SearchFichDefault =True"
    tempstring.WriteLine (Str)
    
    Str = "CustomerOrderDefault =True"
    tempstring.WriteLine (Str)
    
    Str = "CustomerServeplace =0"
    tempstring.WriteLine (Str)
    
    Str = "CustomerSearchDefault =True"
    tempstring.WriteLine (Str)
    
    Str = "CreditCalculate =False"
    tempstring.WriteLine (Str)
    
    Str = "GoodSearchDefault =True"
    tempstring.WriteLine (Str)
    
    Str = "DiscountDefault =0"
    tempstring.WriteLine (Str)
    
    Str = "SrarchInputDelayKeyboard =500"
    tempstring.WriteLine (Str)
    
    Str = "MaxRecordCount =300"
    tempstring.WriteLine (Str)
    
    Str = "FactorSortItems =0"
    tempstring.WriteLine (Str)
    
    Str = "MojodiControlDefault =False"
    tempstring.WriteLine (Str)
    
    Str = "RowMojodiControl =False"
    tempstring.WriteLine (Str)
    
'    Str = "CommandView =True"
'    tempstring.WriteLine (Str)
'
    Str = "EscapeInvoiceFactor =0"
    tempstring.WriteLine (Str)

    Str = "StartUpFormDefault =0"
    tempstring.WriteLine (Str)
    
    Str = "Barcodelengh =16"
    tempstring.WriteLine (Str)
    
    Str = "BarcodeChance =False"
    tempstring.WriteLine (Str)
    
    Str = "PriceChance =50000"
    tempstring.WriteLine (Str)
    
    Str = "RefreshFichNo =False"
    tempstring.WriteLine (Str)
    
    Str = "DirectBascule =False"
    tempstring.WriteLine (Str)
    
    Str = "StopOnEditFich =False"
    tempstring.WriteLine (Str)
    
    Str = "MultiPrice =False"
    tempstring.WriteLine (Str)
    
    Str = "UpdateBuyPrice =False"
    tempstring.WriteLine (Str)
    
    Str = "UpdateSellprice =False"
    tempstring.WriteLine (Str)
    
    Str = "TrazooBarcode =False"
    tempstring.WriteLine (Str)
    
    Str = "CodeFlag =False"
    tempstring.WriteLine (Str)
    
    Str = "SoundAlarm =True"
    tempstring.WriteLine (Str)
    
    Str = "TypeBascule =0"
    tempstring.WriteLine (Str)
    
    Str = "CashPayment =True"
    tempstring.WriteLine (Str)
    
    Str = "InpersonSalonPayment =True"
    tempstring.WriteLine (Str)
    
    Str = "InpersonDeliveryPayment =False"
    tempstring.WriteLine (Str)
    
    Str = "InpersonOutPayment =True"
    tempstring.WriteLine (Str)
    
    Str = "InpersonTablePayment =False"
    tempstring.WriteLine (Str)
    
    Str = "InpersonSalonBalance =True"
    tempstring.WriteLine (Str)
    
    Str = "InpersonDeliveryBalance =False"
    tempstring.WriteLine (Str)
    
    Str = "InpersonOutBalance =True"
    tempstring.WriteLine (Str)
    
    Str = "InpersonTableBalance =False"
    tempstring.WriteLine (Str)
    
    Str = "ByPhoneSalonPayment =True"
    tempstring.WriteLine (Str)
    
    Str = "ByPhoneDeliveryPayment =False"
    tempstring.WriteLine (Str)
    
    Str = "ByPhoneTablePayment =False"
    tempstring.WriteLine (Str)
    
    Str = "ByPhoneSalonBalance =True"
    tempstring.WriteLine (Str)
    
    Str = "ByPhoneDeliveryBalance =False"
    tempstring.WriteLine (Str)
    
    Str = "ByPhoneTableBalance =False"
    tempstring.WriteLine (Str)
    
    
    Str = "ThreeSegmentSearch =False"
    tempstring.WriteLine (Str)
    
    Str = "NumberOfUnitSale =False"
    tempstring.WriteLine (Str)
    
    Str = "MenuViewAfterGood =False"
    tempstring.WriteLine (Str)
    
    Str = "PayFactorView =False"
    tempstring.WriteLine (Str)

    Str = "FromStoreFee =0"
    tempstring.WriteLine (Str)
    
    Str = "BarcodeAutoEscape =False"
    tempstring.WriteLine (Str)

    Str = "PrintAfterTasvieh =False"
    tempstring.WriteLine (Str)

    Str = "AutoBarcode =False"
    tempstring.WriteLine (Str)

    Str = "AutoCallerId =True"
    tempstring.WriteLine (Str)

    Str = "CustomerAscii =False"
    tempstring.WriteLine (Str)

    Str = "CustomerFarsi =False"
    tempstring.WriteLine (Str)
    
    Str = "CustomerOnlinePrice =False"
    tempstring.WriteLine (Str)

    Str = "NoCurrentDay =False"
    tempstring.WriteLine (Str)
    
    Str = "SaleStartDefault =0"
    tempstring.WriteLine (Str)
    
     Str = "ShowDigitNumber =0"
     tempstring.WriteLine (Str)
    
     Str = "CustomerRate =0"
     tempstring.WriteLine (Str)

    Str = "GoodPercentage =False"
    tempstring.WriteLine (Str)

     Str = "CallerIdSpace =1"
     tempstring.WriteLine (Str)
'''
'''     str = "Callwaiting =True"
'''    tempstring.WriteLine (str)
    
    Str = "CountCustomerDailyBuy =0"
    tempstring.WriteLine (Str)
   
    Str = "CountCustomerGood =0"
    tempstring.WriteLine (Str)
    
    Str = "InvoiceRows =7"
    tempstring.WriteLine (Str)
   
    Str = "PurchaseRows =9"
    tempstring.WriteLine (Str)
   
    Str = "FinalCheck =False"
    tempstring.WriteLine (Str)
   
    Str = "AutoTip =False"
    tempstring.WriteLine (Str)
   
     Str = "FichStatusBar =True"
    tempstring.WriteLine (Str)
    
     Str = "CycleStockNoDefault =1"
    tempstring.WriteLine (Str)
    
    Str = "TextIconViewH =True"
    tempstring.WriteLine (Str)
      
    Str = "TextIconViewV =True"
    tempstring.WriteLine (Str)

    Str = "ForceSeller =False"
    tempstring.WriteLine (Str)
    
    Str = "InvoiceStatusDefault =False"
    tempstring.WriteLine (Str)

    Str = "NumberOfId =8"
    tempstring.WriteLine (Str)
    
    Str = "DiscoveryPort =4111"
    tempstring.WriteLine (Str)
    
    Str = "ResponsePort =4112"
    tempstring.WriteLine (Str)
    
    Str = "CityCode =21"
    tempstring.WriteLine (Str)
    
    Str = "AlmLogFile =False"
    tempstring.WriteLine (Str)
    
    Str = "NetworkCallerId =False"
    tempstring.WriteLine (Str)

    Str = "TzWeightBarcodeId =200"
    tempstring.WriteLine (Str)

    Str = "CustomerFeeDataBase =False"
    tempstring.WriteLine (Str)

    Str = "GoodLevelAutoDiscount =False"
    tempstring.WriteLine (Str)
    
    Str = "StartNumberCartReader =1"
    tempstring.WriteLine (Str)
    
    Str = "NumberOfCardReader =5"
    tempstring.WriteLine (Str)
   
    Str = "ReadDirectWeightTz1 =False"
    tempstring.WriteLine (Str)
    
    Str = "TaxView =False"
    tempstring.WriteLine (Str)
    
    Str = "OutPrice =1"
    tempstring.WriteLine (Str)
    
    Str = "TzNumericBarcodeId =300"
    tempstring.WriteLine (Str)
    
    Str = "NumberOfUnitBuy =False"
    tempstring.WriteLine (Str)
    
    Str = "EditCompatibleSamar1 =False"
    tempstring.WriteLine (Str)
    
    Str = "UndoRedoCompatibleSamar1 =False"
    tempstring.WriteLine (Str)
    
    Str = "EscNotExit =False"
    tempstring.WriteLine (Str)
    
    Str = "ViewTempAddress =0"
    tempstring.WriteLine (Str)
    
    Str = "CancelNotExit =False"
    tempstring.WriteLine (Str)
    
    Str = "ResiveFormerFichCurrentDay =False"
    tempstring.WriteLine (Str)
    
    Str = "SelectInventory =False"
    tempstring.WriteLine (Str)
    
    Str = "AlphabetGoodSearch =True"
    tempstring.WriteLine (Str)
    
    Str = "DeficitLog =False"
    tempstring.WriteLine (Str)
    
'    Str = "CallerId8Port =False"
'    tempstring.WriteLine (Str)
'
    Str = "ShiftRate =False"
    tempstring.WriteLine (Str)

    Str = "SellerCaption =›—Ê‘‰œÂ"
    tempstring.WriteLine (Str)
    
    Str = "SelectSeller =False"
    tempstring.WriteLine (Str)
   
    Str = "CountCustomerShiftBuy =0"
    tempstring.WriteLine (Str)
    
    Str = "FixRateChange =False"
    tempstring.WriteLine (Str)
      
    Str = "AutoCashClose =False"
    tempstring.WriteLine (Str)
    
    Str = "PrintAfterDeliver =False"
    tempstring.WriteLine (Str)
    
    Str = "UpDateCarryFee =False"
    tempstring.WriteLine (Str)
    
    Str = "ShowClock =True"
    tempstring.WriteLine (Str)
    
    Str = "VoiceRecord =False"
    tempstring.WriteLine (Str)
    
    Str = "TreeViewMenu =False"
    tempstring.WriteLine (Str)
    
    Str = "CallerIdAutoView =True"
    tempstring.WriteLine (Str)
    
    Str = "CallerIdTest =False"
    tempstring.WriteLine (Str)
    
    Str = "Pager =False"
    tempstring.WriteLine (Str)
    
    Str = "AutoBackup =True"
    tempstring.WriteLine (Str)
    
    Str = "PosModel =1"
    tempstring.WriteLine (Str)
    
    Str = "PosPayment =False"
    tempstring.WriteLine (Str)
    
    Str = "PassPhrase =50000"
    tempstring.WriteLine (Str)

    Str = "PosApprovedText =<Approved>APPROVED</Approved>"
    tempstring.WriteLine (Str)
    
    Str = "ReportHeadername =¬—Ì«"
    tempstring.WriteLine (Str)
    
    Str = "FastCustSave =False"
    tempstring.WriteLine (Str)
    
    Str = "RepetitiveGood =False"
    tempstring.WriteLine (Str)
    
    Str = "OtherPartition =False"
    tempstring.WriteLine (Str)
    
    Str = "NotShowPrintNotice =False"
    tempstring.WriteLine (Str)
    
    Str = "MultiInventory =False"
    tempstring.WriteLine (Str)
    
    Str = "Frame_Printers =False"
    tempstring.WriteLine (Str)
    
    Str = "HasOptionPrice =False"
    tempstring.WriteLine (Str)
    
    Str = "ShowOption =False"
    tempstring.WriteLine (Str)
    
    Str = "LableUsedGood =False"
    tempstring.WriteLine (Str)
    
    Str = "LoyaltyCustomers =False"
    tempstring.WriteLine (Str)
    
    Str = "LoyaltyAllCustomers =False"
    tempstring.WriteLine (Str)
    
    Str = "StartCharacter =%"
    tempstring.WriteLine (Str)
    
    Str = "OneClickShow =False"
    tempstring.WriteLine (Str)
    
    Str = "TouchScreen =False"
    tempstring.WriteLine (Str)
    
    Str = "NoRowMenu =1"
    tempstring.WriteLine (Str)
    
    Str = "RfidReader =False"
    tempstring.WriteLine (Str)
    
    Str = "RfidInterval =2000"
    tempstring.WriteLine (Str)
    
    Str = "RfidLongBuzzer =True"
    tempstring.WriteLine (Str)
    
    Str = "TelNetServerActive =False"
    tempstring.WriteLine (Str)
    
    Str = "TelNetServerIP =192.168.1.1"
    tempstring.WriteLine (Str)
    
    Str = "TelNetServerPort =2001"
    tempstring.WriteLine (Str)
    
    Str = "PrintAfterPayk =False"
    tempstring.WriteLine (Str)

    Str = "PrintAfterOrder =False"
    tempstring.WriteLine (Str)

    Str = "LabelPrint =False"
    tempstring.WriteLine (Str)
    
    Str = "PersonIdCheck =False"
    tempstring.WriteLine (Str)
    
    Str = "PersonIdRefreshTime =5"
    tempstring.WriteLine (Str)
    
    Str = "ListAutoLoad =False"
    tempstring.WriteLine (Str)

    Str = "DeviceIP =192.168.1.100"
    tempstring.WriteLine (Str)
    
    Str = "DeviceID =1"
    tempstring.WriteLine (Str)
    
    Str = "ListFont =14"
    tempstring.WriteLine (Str)
    
    Str = "Device2IP =192.168.1.101"
    tempstring.WriteLine (Str)
    
    Str = "Device2Id =2"
    tempstring.WriteLine (Str)
    
     tempstring.Close

End Function
Public Function SetDefaultInvoiceSettingFile()

    Dim f As New FileSystemObject
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
    
    f.CreateTextFile InvoiceSettingFile
      
    Set tempstring = filetemp.OpenTextFile(InvoiceSettingFile, ForWriting, False, TristateFalse)
    
    Str = "ColRow =1"
    tempstring.WriteLine (Str)
    
    Str = "ColFee =1"
    tempstring.WriteLine (Str)
    
    Str = "ColTotal =1"
    tempstring.WriteLine (Str)
    
    Str = "ColGoodCode =1"
    tempstring.WriteLine (Str)

    Str = "ColChanges =1"
    tempstring.WriteLine (Str)
    
    Str = "ColSeller =0"
    tempstring.WriteLine (Str)
    
    Str = "ColDiscount =0"
    tempstring.WriteLine (Str)
    
    Str = "ColRate =0"
    tempstring.WriteLine (Str)
    
    Str = "ColStore =0"
    tempstring.WriteLine (Str)
    
    Str = "ColMojodi =1"
    tempstring.WriteLine (Str)
    
    Str = "ColUnitGood =1"
    tempstring.WriteLine (Str)
    
    Str = "ColTax =1"
    tempstring.WriteLine (Str)
    
    Str = "ColDuty =1"
    tempstring.WriteLine (Str)
    
    Str = "ShowPictureGood =0"
    tempstring.WriteLine (Str)
    
    Str = "ShowGoodTime =1000"
    tempstring.WriteLine (Str)
    
    Str = "ShowInvoiceMenu =0"
    tempstring.WriteLine (Str)
    
    Str = "GoodMenuView =0"
    tempstring.WriteLine (Str)

    Str = "GoodMenuFileName =Total_GoodMenu"
    tempstring.WriteLine (Str)

    Str = "ScreenSaverTime =0"
    tempstring.WriteLine (Str)

    Str = "LanguageIcon =1"
    tempstring.WriteLine (Str)

    Str = "KeyboardIcon =1"
    tempstring.WriteLine (Str)

    Str = "ColorIcon =1"
    tempstring.WriteLine (Str)

    Str = "TelephoneIcon =1"
    tempstring.WriteLine (Str)

    Str = "ShowLogo =0"
    tempstring.WriteLine (Str)

    Str = "PrintLable =False"
    tempstring.WriteLine (Str)

    Str = "AryaSmsPanel =False"
    tempstring.WriteLine (Str)

    Str = "ForceTax =False"
    tempstring.WriteLine (Str)

    tempstring.Close

End Function
Public Function SetInvoiceSettingFile()

    Dim f As New FileSystemObject
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
     
    Set tempstring = filetemp.OpenTextFile(InvoiceSettingFile, ForWriting, False, TristateFalse)
    
    Str = "ColRow =" & clsInvoiceValue.ColRow
    tempstring.WriteLine (Str)
    
    Str = "ColFee =" & clsInvoiceValue.ColFee
    tempstring.WriteLine (Str)
    
    Str = "ColTotal =" & clsInvoiceValue.ColTotal
    tempstring.WriteLine (Str)
    
    Str = "ColGoodCode =" & clsInvoiceValue.ColGoodCode
    tempstring.WriteLine (Str)

    Str = "ColChanges =" & clsInvoiceValue.ColChanges
    tempstring.WriteLine (Str)
        
    Str = "ColSeller =" & clsInvoiceValue.ColSeller
    tempstring.WriteLine (Str)
    
    Str = "ColDiscount =" & clsInvoiceValue.ColDiscount
    tempstring.WriteLine (Str)
    
    Str = "ColRate =" & clsInvoiceValue.ColRate
    tempstring.WriteLine (Str)
    
    Str = "ColStore =" & clsInvoiceValue.ColStore
    tempstring.WriteLine (Str)
    
    Str = "ColMojodi =" & clsInvoiceValue.ColMojodi
    tempstring.WriteLine (Str)
    
    Str = "ColUnitGood =" & clsInvoiceValue.ColUnitGood
    tempstring.WriteLine (Str)
    
    Str = "ColTax =" & clsInvoiceValue.ColTax
    tempstring.WriteLine (Str)
    
    Str = "ColDuty =" & clsInvoiceValue.ColDuty
    tempstring.WriteLine (Str)
    
    Str = "ShowPictureGood =" & clsInvoiceValue.ShowPictureGood
    tempstring.WriteLine (Str)
    
    Str = "ShowGoodTime =" & clsInvoiceValue.ShowGoodTime
    tempstring.WriteLine (Str)
    
    Str = "ShowInvoiceMenu =" & clsInvoiceValue.ShowInvoiceMenu
    tempstring.WriteLine (Str)
    
    Str = "GoodMenuView =" & clsInvoiceValue.GoodMenuView
    tempstring.WriteLine (Str)
    
    Str = "GoodMenuFileName =" & clsInvoiceValue.GoodMenuFileName
    tempstring.WriteLine (Str)

    Str = "ScreenSaverTime =" & clsInvoiceValue.ScreenSaverTime
    tempstring.WriteLine (Str)

    Str = "LanguageIcon =" & clsInvoiceValue.LanguageIcon
    tempstring.WriteLine (Str)

    Str = "KeyboardIcon =" & clsInvoiceValue.KeyboardIcon
    tempstring.WriteLine (Str)

    Str = "ColorIcon =" & clsInvoiceValue.ColorIcon
    tempstring.WriteLine (Str)

    Str = "TelephoneIcon =" & clsInvoiceValue.TelephoneIcon
    tempstring.WriteLine (Str)

    Str = "ShowLogo =" & clsInvoiceValue.ShowLogo
    tempstring.WriteLine (Str)

    Str = "PrintLable =" & clsInvoiceValue.PrintLable
    tempstring.WriteLine (Str)

    tempstring.Close

End Function
Public Function SetDefaultUserSettingFile()

    Dim f As New FileSystemObject
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
    
    f.CreateTextFile UserSettingFile
      
    Set tempstring = filetemp.OpenTextFile(UserSettingFile, ForWriting, False, TristateFalse)
    
    Str = "Invoice_BackColorForm =12640511"
    tempstring.WriteLine (Str)
    
    Str = "Invoice_BackColorBtn0 =12640511"
    tempstring.WriteLine (Str)

    Str = "Invoice_BackColorBtn1 =12640511"
    tempstring.WriteLine (Str)

    Str = "Invoice_BackColorBtn2 =12640511"
    tempstring.WriteLine (Str)

    Str = "Invoice_BackColorBtn3 =12640511"
    tempstring.WriteLine (Str)

    Str = "Invoice_BackColorBtn4 =12640511"
    tempstring.WriteLine (Str)

    Str = "Invoice_BackColorBtn5 =12640511"
    tempstring.WriteLine (Str)

    Str = "Invoice_BackColorBtn6 =12640511"
    tempstring.WriteLine (Str)

    Str = "Invoice_BackColorBtn7 =12640511"
    tempstring.WriteLine (Str)

    Str = "Invoice_BackColorFlexGrid =12640511"
    tempstring.WriteLine (Str)
    
    Str = "Purchase_BackColorForm = 16761024"
    tempstring.WriteLine (Str)
    
    Str = "Purchase_BackColorBtn = 11829623"
    tempstring.WriteLine (Str)
    
    Str = "Purchase_BackColorFlexGrid = 16744576"
    tempstring.WriteLine (Str)
    
    Str = "Invoice_FontMenuName =Tahoma"
    tempstring.WriteLine (Str)
    
    Str = "Invoice_FontMenuSize =11"
    tempstring.WriteLine (Str)
    
    Str = "Invoice_FontMenuBold =False"
    tempstring.WriteLine (Str)
    
    Str = "Invoice_FontFlexGridName =Nazanin"
    tempstring.WriteLine (Str)
    
    Str = "Invoice_FontFlexGridSize =12"
    tempstring.WriteLine (Str)
    
    Str = "Invoice_FontFlexGridBold =True"
    tempstring.WriteLine (Str)
 
    Str = "Invoice_FontDifferencesName =Nazanin"
    tempstring.WriteLine (Str)
    
    Str = "Invoice_FontDifferencesSize =14"
    tempstring.WriteLine (Str)
    
    Str = "Invoice_FontDifferencesBold =True"
    tempstring.WriteLine (Str)
   
 tempstring.Close

End Function
Public Function setDefaultAryaSettingFile()

    Dim f As New FileSystemObject
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
    
    f.CreateTextFile AryaSettingFile
      
    Set tempstring = filetemp.OpenTextFile(AryaSettingFile, ForWriting, False, TristateFalse)
    Str = "ServerName =."
    tempstring.WriteLine (Str)
    
    Str = "DbName =Total"
    tempstring.WriteLine (Str)
    
    Str = "StationName = F.G.Arya Co"
    tempstring.WriteLine (Str)
    
    Str = "StationNo =1"
    tempstring.WriteLine (Str)
    
    Str = "AppPath =D:\Arya"
    tempstring.WriteLine (Str)

    Str = "ExternalDataName =Total_Ext"
    tempstring.WriteLine (Str)
    
    Str = "ExternalDbPath =L:\Data\Total_Ext.mdf"
    tempstring.WriteLine (Str)

    Str = "RestoreDataBaseData =D:\Arya\Data"
    tempstring.WriteLine (Str)
    
    Str = "RestoreDataBaseLog =D:\Arya\Data"
    tempstring.WriteLine (Str)

    Str = "AccSrvName =."
    tempstring.WriteLine (Str)
    
    Str = "AccdataBaseName =Account"
    tempstring.WriteLine (Str)
    
    Str = "CustomerDisplayName =Arya"
    tempstring.WriteLine (Str)
    
    Str = "AccountSystemName =Samar"
    tempstring.WriteLine (Str)
    
'    Str = "EnableUpperAmountGood =0"
'    tempstring.WriteLine (Str)
'
    Str = "Externalaccounting =False"
    tempstring.WriteLine (Str)
    
    Str = "BranchView =True"
    tempstring.WriteLine (Str)
    
    Str = "SecVersion =0"
    tempstring.WriteLine (Str)
    
    Str = "HVersion =1"
    tempstring.WriteLine (Str)
    
    Str = "PrintServer =False"
    tempstring.WriteLine (Str)
    
    Str = "SoftLock =False"
    tempstring.WriteLine (Str)
    
    Str = "MiladiDate =0"
    tempstring.WriteLine (Str)
    
    Str = "DBLogin =sa"
    tempstring.WriteLine (Str)
    
    Str = "NetLock =False"
    tempstring.WriteLine (Str)
    
    Str = "CustomerName =¬—Ì«"
    tempstring.WriteLine (Str)
    
    Str = "CustomerAddress =«Ì—«‰"
    tempstring.WriteLine (Str)
    
    Str = "UnitPrice =—Ì«·"
    tempstring.WriteLine (Str)

    Str = "DBPass =lemon7430"
    tempstring.WriteLine (Str)

    Str = "NewPrinting =False"
    tempstring.WriteLine (Str)


    tempstring.Close

End Function

Public Function setAryaSettingFile()

    Dim f As New FileSystemObject
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
    
      
    Set tempstring = filetemp.OpenTextFile(AryaSettingFile, ForWriting, False, TristateFalse)
    Str = "ServerName =" & clsArya.ServerName
    tempstring.WriteLine (Str)
    
    Str = "DbName =" & clsArya.DbName
    tempstring.WriteLine (Str)
    
    Str = "StationName =" & clsArya.StationName
    tempstring.WriteLine (Str)
    
    Str = "StationNo =" & clsArya.StationNo
    tempstring.WriteLine (Str)
     
    Str = "AppPath =" & clsArya.AppPath
    tempstring.WriteLine (Str)

    Str = "ExternalDataName =" & clsArya.ExternalDataName
    tempstring.WriteLine (Str)
    
    Str = "ExternalDbPath =" & clsArya.ExternalDbPath
    tempstring.WriteLine (Str)

    Str = "RestoreDataBaseData =D:\Arya\Data"
    tempstring.WriteLine (Str)
    
    Str = "RestoreDataBaseLog =D:\Arya\Data"
    tempstring.WriteLine (Str)

    Str = "AccSrvName =" & clsArya.AccSrvName
    tempstring.WriteLine (Str)
    
    Str = "AccdataBaseName =" & clsArya.AccdataBaseName
    tempstring.WriteLine (Str)
    
    Str = "CustomerDisplayName =" & clsArya.CustomerDisplayName
    tempstring.WriteLine (Str)
    
    Str = "AccountSystemName =" & clsArya.AccountSystemName
    tempstring.WriteLine (Str)
    
'    Str = "EnableUpperAmountGood =" & clsArya.EnableUpperAmountGood
'    tempstring.WriteLine (Str)
'
    Str = "ExternalAccounting =" & clsArya.ExternalAccounting
    tempstring.WriteLine (Str)
    
    Str = "BranchView =" & clsArya.BranchView
    tempstring.WriteLine (Str)
    
    Str = "SecVersion =" & clsArya.SecVersion
    tempstring.WriteLine (Str)
    
    Str = "HVersion =" & clsArya.HVersion
    tempstring.WriteLine (Str)
    
    Str = "PrintServer =" & clsArya.PrintServer
    tempstring.WriteLine (Str)
    
    Str = "SoftLock =" & clsArya.SoftLock
    tempstring.WriteLine (Str)
    
    Str = "MiladiDate =" & clsArya.MiladiDate
    tempstring.WriteLine (Str)
    
    Str = "DBLogin =" & clsArya.DBLogin
    tempstring.WriteLine (Str)
    
    Str = "NetLock =" & clsArya.NetLock
    tempstring.WriteLine (Str)
    
    Str = "CustomerName =" & Trim(clsArya.Company)
    tempstring.WriteLine (Str)
    
    Str = "CustomerAddress =" & Trim(clsArya.CustomerAddres)
    tempstring.WriteLine (Str)
    
    Str = "UnitPrice =" & Trim(clsArya.UnitPrice)
    tempstring.WriteLine (Str)
    
    Str = "DBPass =" & clsArya.DBPass
    tempstring.WriteLine (Str)
    
    Str = "NewPrinting =" & clsArya.NewPrinting
    tempstring.WriteLine (Str)
    
    tempstring.Close

End Function

Public Function setAccountingSettingFile()
    Dim file As New FileSystemObject
    Dim strWriter As TextStream
    Dim strTemp As String
    Dim indexEqual As Integer
    
    Set strWriter = file.OpenTextFile(AccountingSettingFile, ForWriting, False, TristateFalse)
    strTemp = "PerssonelAtf =" & clsAccounting.PerssonelAtf
    strWriter.WriteLine (strTemp)
    
    strTemp = "CustomerAtf =" & clsAccounting.CustomerAtf
    strWriter.WriteLine (strTemp)
    
    strTemp = "SupplierAtf =" & clsAccounting.SupplierAtf
    strWriter.WriteLine (strTemp)
    
    strTemp = "ConsumerCompanyAtf =" & clsAccounting.CosumerCompanyAtf
    strWriter.WriteLine (strTemp)
    
    strTemp = "SupplierCompanyAtf =" & clsAccounting.SupplierCompanyAtf
    strWriter.WriteLine (strTemp)
    
    strWriter.Close
    
End Function
Public Function setDefaultAccountingSettingFile()
    Dim file As New FileSystemObject
    Dim strWriter As TextStream
    Dim strTemp As String
    Dim indexEqual As Integer
    
    Set strWriter = file.OpenTextFile(AccountingSettingFile, ForWriting, True, TristateFalse)
    strTemp = "PerssonelAtf =2"
    strWriter.WriteLine (strTemp)
    
    strTemp = "CustomerAtf =4"
    strWriter.WriteLine (strTemp)
    
    strTemp = "SupplierAtf =5"
    strWriter.WriteLine (strTemp)
    
    strTemp = "ConsumerCompanyAtf =5"
    strWriter.WriteLine (strTemp)
    
    strTemp = "SupplierCompanyAtf =5"
    strWriter.WriteLine (strTemp)
    
    strWriter.Close
End Function

Public Function SetDefaultGoodMenuSettingFile()

    Dim f As New FileSystemObject
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
    Dim i  As Long
    
    f.CreateTextFile GoodMenuSettingFile
      
    Set tempstring = filetemp.OpenTextFile(GoodMenuSettingFile, ForWriting, False, TristateFalse)
    
    For i = 0 To 5
        Str = "ViewSegmant" & i & " =True"
        tempstring.WriteLine (Str)
    
        Str = "HeaderTitr" & i & " =" & "«’·Ì " & CStr(i + 1)
        tempstring.WriteLine (Str)
    
        Str = "HeaderFont" & i & " =Titr"
        tempstring.WriteLine (Str)
    
        Str = "HeaderSizeFont" & i & " =14"
        tempstring.WriteLine (Str)
    
        Str = "HeaderColorFont" & i & " =&H00000080&"
        tempstring.WriteLine (Str)
    
        Str = "GridFont" & i & " =Nazanin"
        tempstring.WriteLine (Str)
    
        Str = "GridSizeFont" & i & " =14"
        tempstring.WriteLine (Str)
    
        Str = "GridColorFont" & i & " =&H80000008&"
        tempstring.WriteLine (Str)
    
        Str = "ViewRow" & i & " =True"
        tempstring.WriteLine (Str)
    
        Str = "ViewName" & i & " =True"
        tempstring.WriteLine (Str)
    
        Str = "ViewFee1" & i & " =True"
        tempstring.WriteLine (Str)
    
        Str = "ViewFee2" & i & " =True"
        tempstring.WriteLine (Str)
    
        Str = "ViewPicture" & i & " =True"
        tempstring.WriteLine (Str)
    
        Str = "ViewDescription" & i & " =True"
        tempstring.WriteLine (Str)
    
        Str = "RowName" & i & " =" & "—œÌ›"
        tempstring.WriteLine (Str)
    
        Str = "GoodName" & i & " =" & "‰«„ ò«·«"
        tempstring.WriteLine (Str)
    
        Str = "Fee1Name" & i & " =" & "›Ì ”«·‰"
        tempstring.WriteLine (Str)
    
        Str = "Fee2Name" & i & " =" & "›Ì »Ì—Ê‰"
        tempstring.WriteLine (Str)
    
    Next i
    tempstring.Close
    
End Function
Public Function SetGoodMenuSettingFile()

    Dim f As New FileSystemObject
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
    Dim i  As Long
    Set tempstring = filetemp.OpenTextFile(GoodMenuSettingFile, ForWriting, False, TristateFalse)
    
    
    For i = 0 To 5
        Str = "ViewSegmant" & i & " =" & clsGoodMenu.ViewSegmant(i)
        tempstring.WriteLine (Str)
    
        Str = "HeaderTitr" & i & " =" & clsGoodMenu.HeaderTitr(i)
        tempstring.WriteLine (Str)
    
        Str = "HeaderFont" & i & " =" & clsGoodMenu.HeaderFont(i)
        tempstring.WriteLine (Str)
    
        Str = "HeaderSizeFont" & i & " =" & clsGoodMenu.HeaderSizeFont(i)
        tempstring.WriteLine (Str)
    
        Str = "HeaderColorFont" & i & " =" & clsGoodMenu.HeaderColorFont(i)
        tempstring.WriteLine (Str)
    
        Str = "GridFont" & i & " =" & clsGoodMenu.GridFont(i)
        tempstring.WriteLine (Str)
    
        Str = "GridSizeFont" & i & " =" & clsGoodMenu.GridSizeFont(i)
        tempstring.WriteLine (Str)
    
        Str = "GridColorFont" & i & " =" & clsGoodMenu.GridColorFont(i)
        tempstring.WriteLine (Str)
    
        Str = "ViewRow" & i & " =" & clsGoodMenu.ViewRow(i)
        tempstring.WriteLine (Str)
    
        Str = "ViewName" & i & " =" & clsGoodMenu.ViewName(i)
        tempstring.WriteLine (Str)
    
        Str = "ViewFee1" & i & " =" & clsGoodMenu.ViewFee1(i)
        tempstring.WriteLine (Str)
    
        Str = "ViewFee2" & i & " =" & clsGoodMenu.ViewFee2(i)
        tempstring.WriteLine (Str)
    
        Str = "ViewPicture" & i & " =" & clsGoodMenu.ViewPicture(i)
        tempstring.WriteLine (Str)
    
        Str = "ViewDescription" & i & " =" & clsGoodMenu.ViewDescription(i)
        tempstring.WriteLine (Str)
    
        Str = "RowName" & i & " =" & clsGoodMenu.RowName(i)
        tempstring.WriteLine (Str)
    
        Str = "GoodName" & i & " =" & clsGoodMenu.GoodName(i)
        tempstring.WriteLine (Str)
    
        Str = "Fee1Name" & i & " =" & clsGoodMenu.Fee1Name(i)
        tempstring.WriteLine (Str)
    
        Str = "Fee2Name" & i & " =" & clsGoodMenu.Fee2Name(i)
        tempstring.WriteLine (Str)
    
    Next i
    tempstring.Close
    
End Function

