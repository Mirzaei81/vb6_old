VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Printing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Dim Parameter() As Parameter
Dim ClsCnn As New ADODB.Connection
Dim HostName As String
Dim PortName As String
Dim fs As New FileSystemObject
Dim PartitionNo As Long
Dim Islog As Boolean
Dim IsPing As Boolean
Dim PingPosition, LanAddress As String

Public Function Printing(FichNumber As Long, intStationId As Integer, ByVal AddEditMode As EnumAddEditMode, ActionMode As EnumActionLog) As Boolean  ', basNo As Integer , mystr As String
    On Error GoTo ErrorHandler
    
    Dim RstTemp As New ADODB.Recordset
    Dim RstTemp2 As New ADODB.Recordset
    Dim j As Long
    If ActionMode = 0 Then
        ActionMode = EnumActionLog.Printing
    End If
    
    Unload frmFactorReceived
    Printing = True
    
    Dim ParametersTmp(3) As Parameter
    ParametersTmp(0) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    ParametersTmp(1) = GenerateInputParameter("@No", adBigInt, 8, FichNumber)
    ParametersTmp(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    ParametersTmp(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set RstTemp2 = RunParametricStoredProcedure2Rec("GetFacMinfo", ParametersTmp, ClsCnn)

    If Not (RstTemp2.EOF = True And RstTemp2.BOF = True) Then
        If RstTemp2!PartitionID > 0 Then
            PartitionNo = RstTemp2!PartitionID
        Else
            PartitionNo = clsStation.PartitionID
        End If
        Dim Parameters(5) As Parameter
        Parameters(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameters(1) = GenerateInputParameter("@ServePlace", adInteger, 4, RstTemp2!ServePlace)
        Parameters(2) = GenerateInputParameter("@CurrentStationId", adInteger, 4, intStationId)
        Parameters(3) = GenerateInputParameter("@PartitionId", adInteger, 4, PartitionNo)
        Parameters(4) = GenerateInputParameter("@Mode", adInteger, 4, AddEditMode)
        Parameters(5) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
        If RstTemp.State <> 0 Then RstTemp.Close
        Set RstTemp = RunParametricStoredProcedure2Rec("GetPrintInfo", Parameters, ClsCnn)
        
        Dim ReportFileName  As String
        Dim LowLevelPortName As String
        Dim i As Integer
            
        If Not (RstTemp.BOF Or RstTemp.EOF) Then
            While RstTemp.EOF = False
                If (RstTemp.Fields("PermittedModes").Value And AddEditMode) = AddEditMode Then
                    
                    ReportFileName = App.Path & "\Reports" & RepVer & "\" & RstTemp!rptFilePath
                    If fs.FileExists(ReportFileName) = False Then
                        ShowDisMessage "›«Ì· »—«Ì ç«Å Ì«›  ‰‘œ ", 2000
                        MsgBox "«Ì‰ ›«Ì· Ì«›  ‰‘œ: " & vbCrLf & ReportFileName
                        Printing = False
                    Else
                        Islog = RstTemp.Fields("Islog").Value
                        IsPing = RstTemp.Fields("IsPing").Value
                        LanAddress = IIf(IsNull(RstTemp.Fields("LanAddress").Value), "", RstTemp.Fields("LanAddress").Value)
                        HostName = IIf(IsNull(RstTemp.Fields("HostName").Value), "", RstTemp.Fields("HostName").Value)
                        PortName = IIf(IsNull(RstTemp.Fields("Port").Value), "", RstTemp.Fields("Port").Value)
                        If clsArya.PrintServer = True And Islog = True Then
                            Printing = InsertPrintFich(FichNumber, RstTemp!PrintFormat, intStationId, RstTemp!PrinterNo, AddEditMode, PartitionNo)       ' PRINT
                        Else
                            Printing = Doprint(RstTemp!DirectRpt, RstTemp!PrinterName, ReportFileName, RstTemp!PrintFormat, RstTemp!RepeatNo, RstTemp!PrinterNo, FichNumber, RstTemp!StationId, AddEditMode)   ' PRINT
                        End If
                    End If
                End If
                
                RstTemp.MoveNext
            Wend
        Else
            ShowDisMessage " „Ê—œ ç«Å  ⁄—Ì› ‰‘œÂ «”   ", 2000
            Printing = False
        End If
    End If
    Set RstTemp = Nothing
    Set RstTemp2 = Nothing
    If Printing = False Then
        If clsStation.NotShowPrintNotice = False Then
            frmDisMsg.lblMessage = " ç«Å «‰Ã«„ ‰‘œ"
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If
        Exit Function
    End If
    
    Dim intSerialNo As Long
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, FichNumber)
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(2) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
    Parameter(3) = GenerateInputParameter("@ActionCode", adInteger, 4, ActionMode)
    Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    intSerialNo = RunParametricStoredProcedure("InsertHistory", Parameter, ClsCnn)
            
    Exit Function

ErrorHandler:
    LogSaveNew "Printing => ", err.Description, err.Number, err.Source, "Printing"
    ShowDisMessage "œ— ⁄„·Ì«  ç«Å Œÿ«ÌÌ »Â ÊÃÊœ ¬„œÂ «”  ", 2000
    Printing = False
    Set RstTemp = Nothing
    Set RstTemp2 = Nothing
End Function


Public Function Doprint(DirectRpt As Integer, PassedPrinterName As String, ReportFileName As String, PrintFormat As Integer, intNumberOfCopies As Integer, PrinterNo As Integer, Optional FichNumber As Long, Optional intStationId As Integer, Optional AddEditMode As EnumAddEditMode) As Boolean
    
    Dim i As Long
    Dim MyRecordSet As New ADODB.Recordset
    Dim Rst As New ADODB.Recordset
    Dim rctmp As New ADODB.Recordset

    Dim MyCrystalReport
    Set MyCrystalReport = CreateObject("Crystal.CrystalReport")
    
    MyCrystalReport.ReportFileName = ReportFileName
    Doprint = True
    Dim prnPrinter
    Dim IsPrinterNameOk As Boolean
    For Each prnPrinter In Printers
        If InStr(1, prnPrinter.DeviceName, PassedPrinterName, 1) Then
            Set Printer = prnPrinter
            IsPrinterNameOk = True
            Exit For
        End If
    Next
    If IsPrinterNameOk = False Then
        MsgBox PassedPrinterName & "œ—·Ì”  Å—Ì‰ —Â« „ÊÃÊœ ‰Ì” "
''        frmInvoice.FlxDetail.SetFocus
        Exit Function
    End If
    
    ReDim Parameter(7) As Parameter
    Parameter(0) = GenerateInputParameter("@intFacNO", adBigInt, 8, FichNumber)
    Parameter(1) = GenerateInputParameter("@PrintFormat", adInteger, 4, PrintFormat)
    Parameter(2) = GenerateInputParameter("@StationID", adInteger, 4, intStationId)
    Parameter(3) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(4) = GenerateInputParameter("@intPrinterNo", adInteger, 4, PrinterNo)
    Parameter(5) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(6) = GenerateInputParameter("@Mode", adInteger, 4, AddEditMode)
    Parameter(7) = GenerateInputParameter("@PartitionId", adInteger, 4, PartitionNo)
    Set MyRecordSet = RunParametricStoredProcedure2Rec("Get_InvoiceInfo_IsFilter", Parameter, ClsCnn)
    
    If (MyRecordSet.EOF = True And MyRecordSet.BOF = True) Then
        If clsStation.NotShowPrintNotice = False Then
            frmDisMsg.lblMessage = "ò«·« Â«Ì «‰ Œ«» ‘œÂ œ— «Ì‰ ›—„  ç«Å  ⁄—Ì› ‰‘œÂ «”  "
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Timer1.Interval = 500
            frmDisMsg.Show vbModal
        End If
        Exit Function
    Else
        
        Dim filetemp As New FileSystemObject
        Dim tempstring As TextStream
        Dim Str As String
        Dim LenghStr As Integer
        Dim TxtReportFile As String
        Dim IsFileExist As Boolean
        Doprint = True
        
        Dim cnn As New ADODB.Connection
        Dim Rctemp As New Recordset
        Dim CountRecord As Long
        Dim strtemporary As String
        
        Select Case DirectRpt

            Case 0    ' Crystal Output
                
                MyCrystalReport.ReportFileName = ReportFileName
                MyCrystalReport.PrinterDriver = Printer.DriverName
                MyCrystalReport.PrinterName = Printer.DeviceName
                MyCrystalReport.PrinterPort = Printer.Port
                MyCrystalReport.ReportTitle = clsStation.ReportHeadername
                
                MyCrystalReport.PrinterCopies = intNumberOfCopies
                'MyCrystalReport.Destination = crptowindow    'crptowindow
                MyCrystalReport.Destination = crptToPrinter
                 
                If AddEditMode <> Perfrage Then
                    MyCrystalReport.ParameterFields(0) = "@intLanguage" & ";" & clsStation.Language & ";" & "true"
                    MyCrystalReport.ParameterFields(1) = "@intFacNo" & ";" & FichNumber & ";" & "true"
                    MyCrystalReport.ParameterFields(2) = "@PrintFormat" & ";" & PrintFormat & ";" & "true"
                    MyCrystalReport.ParameterFields(3) = "@StationId" & ";" & intStationId & ";" & "true"
                    MyCrystalReport.ParameterFields(4) = "@Status" & ";" & mvarStatus & ";" & "true"
                    MyCrystalReport.ParameterFields(5) = "@intPrinterNo" & ";" & PrinterNo & ";" & "true"
                    MyCrystalReport.ParameterFields(6) = "@Mode" & ";" & AddEditMode & ";" & "true"
                    MyCrystalReport.ParameterFields(7) = "@AccountYear" & ";" & AccountYear & ";" & "true"
                    MyCrystalReport.ParameterFields(8) = "@PartitionId" & ";" & PartitionNo & ";" & "true"
                    MyCrystalReport.ParameterFields(9) = "@Branch" & ";" & CurrentBranch & ";" & "true"
                     If Islog = True Then
                         Dim Pk_Id As Long
                         ReDim Parameter(12) As Parameter
                         Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 8, clsStation.Language)
                         Parameter(1) = GenerateInputParameter("@intFacNo", adInteger, 4, FichNumber)
                         Parameter(2) = GenerateInputParameter("@PrintFormat", adInteger, 4, PrintFormat)
                         Parameter(3) = GenerateInputParameter("@StationId", adInteger, 4, intStationId)
                         Parameter(4) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                         Parameter(5) = GenerateInputParameter("@intPrinterNo", adInteger, 4, PrinterNo)
                         Parameter(6) = GenerateInputParameter("@Mode", adInteger, 4, AddEditMode)
                         Parameter(7) = GenerateInputParameter("@AccountYear", adInteger, 4, AccountYear)
                         Parameter(8) = GenerateInputParameter("@PartitionId", adInteger, 4, PartitionNo)
                         Parameter(9) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                         Parameter(10) = GenerateInputParameter("@PrintPosition", adBoolean, 1, 0)
                         Parameter(11) = GenerateInputParameter("@PingStatus", adVarChar, 255, "")
                         Parameter(12) = GenerateOutputParameter("@Result", adInteger, 4)
                         Pk_Id = RunParametricStoredProcedure("Insert_tPrinters_Log", Parameter, ClsCnn)
                      
                     End If
                     MyCrystalReport.RetrieveDataFiles
                     
                     MyCrystalReport.ProgressDialog = False
                     'MyCrystalReport.DiscardSavedData = True
                     ODBCSetting clsArya.ServerName, clsArya.DbName
                      MyCrystalReport.Connect = CrystallConnection
                      
                      Dim Result As Long
                      Result = MyCrystalReport.PrintReport
                      If Result <> 0 Then
                         ShowMessage "  Œÿ« œ— ç«Å  " & vbLf & MyCrystalReport.LastErrorString, True, False, "ﬁ»Ê·", ""
                         LogSaveNew " Œÿ« œ— ç«Å  => ", MyCrystalReport.LastErrorString, err.Number, err.Source, "Doprint"
                         LogSaveNew " Œÿ« œ— ç«Å  => ", err.Description, err.Number, err.Source, "Doprint"
                      End If
                    ' MyCrystalReport.Action = 1
                Else    'Perfrage (Label)
                
                    ReDim Parameter(4)
                    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, FichNumber)
                    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                    Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                    Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                    
                    Set Rst = RunParametricStoredProcedure2Rec("Get_FacMD_Good", Parameter)
                    
                    Dim ii, kk, jj, MaxCount, UsedCount As Long
                    Dim strGood, strDescription As String
                    ii = 0
                    If Not (Rst.BOF Or Rst.EOF) Then
                        Do While Not (Rst.EOF)
                                
                            cnn.Open strConnectionString
                           ' strtemporary = "Select Count(*) as CountRecord from tGood Where Code = " & Rst!GoodCode & " AND Code NOT IN (SELECT GoodCode  FROM tPrinterGood WHERE intPrinterFormat = " & PrintFormat & ")  "
                            strtemporary = "Select Count(*) as CountRecord from tPrinterGood Where GoodCode = " & Rst!GoodCode & " AND  intPrinterFormat = " & PrintFormat
                            Rctemp.Open strtemporary, cnn, adOpenDynamic, adLockOptimistic, adCmdText
                            If Not (Rctemp.EOF = True And Rctemp.BOF = True) Then
                               CountRecord = Val(Rctemp!CountRecord)
                               If CountRecord = 0 Then
                            
                                    MaxCount = Rst!SumAmount
                                    For kk = 1 To Rst!amount
                                        If clsStation.LableUsedGood = False Then
                                            ii = ii + 1
                                            strGood = Rst!nvcName & " " & IIf(IsNull(Rst!DifferencesDescription), "", Rst!DifferencesDescription)
                                            strDescription = ii & " of " & MaxCount
                                        
                                            MyCrystalReport.ParameterFields(0) = "@FichNo" & ";" & FichNumber & ";" & "true"
                                            MyCrystalReport.ParameterFields(1) = "@strGood" & ";" & strGood & ";" & "true"
                                            MyCrystalReport.ParameterFields(2) = "@strDescription" & ";" & strDescription & ";" & "true"
                                            
                                        
                                            ODBCSetting clsArya.ServerName, clsArya.DbName
                                            MyCrystalReport.Connect = CrystallConnection
                                            MyCrystalReport.RetrieveDataFiles
                                            MyCrystalReport.ProgressDialog = False
                                            MyCrystalReport.Action = 1
                                        
                                        Else
                                            ReDim Parameter(0)
                                            Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, Rst!GoodCode)
                                            Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Used", Parameter)
                                            If Not (rctmp.BOF Or rctmp.EOF) Then
                                                jj = 0
                                                ii = ii + 1
                                                UsedCount = rctmp!UsedCount
                                                Do While Not (rctmp.EOF)
                                                    jj = jj + 1
                                                    strGood = rctmp!nvcName
                                                    ''strDescription = ii & " of " & MaxCount & "-" & jj & " of " & UsedCount
                                                    strDescription = jj & " of " & UsedCount
                                                    MyCrystalReport.ParameterFields(0) = "@FichNo" & ";" & FichNumber & ";" & "true"
                                                    MyCrystalReport.ParameterFields(1) = "@strGood" & ";" & strGood & ";" & "true"
                                                    MyCrystalReport.ParameterFields(2) = "@strDescription" & ";" & strDescription & ";" & "true"
                                                    
                                                
                                                    ODBCSetting clsArya.ServerName, clsArya.DbName
                                                    MyCrystalReport.Connect = CrystallConnection
                                                    MyCrystalReport.RetrieveDataFiles
                                                    MyCrystalReport.ProgressDialog = False
                                                    MyCrystalReport.Action = 1
                                                    rctmp.MoveNext
                                                Loop
                                            Else        '''ç«Å Â„«‰ ò«·«Ì «’·Ì
                                                ii = ii + 1
                                                strGood = Rst!nvcName & " " & IIf(IsNull(Rst!DifferencesDescription), "", Rst!DifferencesDescription)
                                                strDescription = ii & " of " & MaxCount
                                                MyCrystalReport.ParameterFields(0) = "@FichNo" & ";" & FichNumber & ";" & "true"
                                                MyCrystalReport.ParameterFields(1) = "@strGood" & ";" & strGood & ";" & "true"
                                                MyCrystalReport.ParameterFields(2) = "@strDescription" & ";" & strDescription & ";" & "true"
                                                
                                            
                                                ODBCSetting clsArya.ServerName, clsArya.DbName
                                                MyCrystalReport.Connect = CrystallConnection
                                                MyCrystalReport.RetrieveDataFiles
                                                MyCrystalReport.ProgressDialog = False
                                                MyCrystalReport.Action = 1
                                            End If
                                        End If
                                    Next
                            
                               End If
                            End If
                            Rctemp.Close
                            cnn.Close
                            Set cnn = Nothing
                            Rst.MoveNext
                        Loop
                    End If
                
                End If
            
            Case 1      ' Object Printer
                
                   ReDim Parameter(9) As Parameter
                   Parameter(0) = GenerateInputParameter("@intLanguage", adBoolean, 1, clsStation.Language)
                   Parameter(1) = GenerateInputParameter("@intFacNO", adBigInt, 8, FichNumber)
                   Parameter(2) = GenerateInputParameter("@PrintFormat", adInteger, 4, PrintFormat)
                   Parameter(3) = GenerateInputParameter("@StationID", adInteger, 4, intStationId)
                   Parameter(4) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                   Parameter(5) = GenerateInputParameter("@intPrinterNo", adInteger, 4, PrinterNo)
                   Parameter(6) = GenerateInputParameter("@Mode", adInteger, 4, AddEditMode)
                   Parameter(7) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                   Parameter(8) = GenerateInputParameter("@PartitionId", adInteger, 4, PartitionNo)
                   Parameter(9) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                   Set MyRecordSet = RunParametricStoredProcedure2Rec("Get_InvoiceInfo", Parameter)
                   
        '          LogSave "Printing => ", CStr(FichNumber), " ", Time, " "
                   Printer.RightToLeft = True
                   Printer.Font.Bold = True
                   Printer.Font.size = 12
                 '  Printer.FontName = "Trrafic"  ' FontA1x1   Trrafic
                   Printer.FontName = "Nazanin"  ' FontA1x1   Trrafic
                   
                    If filetemp.FileExists(App.Path & "\IMAGE\logo.gif") = True Then
                        Printer.PaintPicture LoadPicture(App.Path & "\IMAGE\logo.gif"), 100, 0, 1000, 1000
                        Printer.Print
                    ElseIf filetemp.FileExists(App.Path & "\IMAGE\logo.jpg") = True Then
                         Printer.PaintPicture LoadPicture(App.Path & "\IMAGE\logo.jpg"), 0, 100, 1000, 1000
                        Printer.Print
                    End If
                   
                   Str = clsStation.ReportHeadername
                   Printer.Print Space((55 - Len(Str)) / 2) & Str
                   Printer.FontName = "B Homa"  ' FontA1x1   Trrafic
                   Str = Space(25) & "’‰œÊﬁ " & intStationId
                   Printer.Print Str '& vbCrLf
                   
                    Printer.Font.size = 12
                    If Right(CStr(FichNumber), 3) = "000" Then Str = "1000" Else Str = Right(CStr(FichNumber), 3)
                    'str = " " & tempstring.ReadLine & ":" & Right(CStr(FichNumber), 3) & "        " & tempstring.ReadLine & ":" & MyRecordSet!Date
                    Str = "   ›Ì‘ : " & Str & "        " & "          «—ÌŒ : " & MyRecordSet!Date
                    Printer.Print Str '& vbCrLf
                    Str = "  ò«—»—     " & CStr(MyRecordSet!User) & "               " & "       ”«⁄  :   " & MyRecordSet!time
                    Printer.Print Str '& vbCrLf
                   
                   Printer.Font.size = 10
                   Printer.Print "----------------------------------------------------------------"  'LF
                   
                    Printer.FontName = "Nazanin"  ' FontA1x1   Trrafic
                    Printer.Font.size = 12
                    If MyRecordSet!Customer > 0 Then
                         Printer.Print "  òœ «‘ —«ò : " & MyRecordSet!MembershipId & "  " & "  ·›‰ : " & Val(MyRecordSet!Tel1)
                         Printer.Print "  „‘ —ò : " & MyRecordSet!Family
                         Printer.Print "  ¬œ—” : " & MyRecordSet!CustomerAddress
                         Printer.Print "----------------------------------------------------------------"  'LF
                    End If
                    Dim Totalprice As Long
                    Totalprice = 0
                   If Not (MyRecordSet.BOF Or MyRecordSet.EOF) Then
                        'Printer.Print tempstring.ReadLine
                        Printer.Print "   ‰«„ ﬂ«·«             „ﬁœ«—     ›Ì           ﬁÌ„  "
'                        Printer.Print "   „ﬁœ«—           ‰«„ ﬂ«·«               ›Ì            Ã„⁄"
                        Printer.Print "----------------------------------------------------------------"  'LF
                      
'                         Printer.CurrentX = Printer.CurrentX
'                         Printer.CurrentX = i
                        Printer.FontName = "Nazanin"  ' FontA1x1   Trrafic
                        Printer.Font.size = 11
                        Do While Not (MyRecordSet.EOF)
                            Totalprice = Totalprice + MyRecordSet!FeeTotal
                            Str = left(Trim(MyRecordSet!GoodName), 15)
                            Printer.Print Str; Tab(16); Space((5 - Len(MyRecordSet!amount)) / 2) & MyRecordSet!amount; Tab(21); MyRecordSet!FeeUnit; Tab(32); MyRecordSet!FeeTotal
                            'Printer.Print Space((5 - Len(MyRecordSet!amount)) / 2) & MyRecordSet!amount & Space((5 - Len(MyRecordSet!amount)) / 2) & Str & Space(22 - Len(Str)) & MyRecordSet!FeeUnit & Space(8 - Len(MyRecordSet!FeeUnit)) & MyRecordSet!FeeTotal
                            'Printer.Print Tab((5 - Len(MyRecordSet!amount)) / 2); MyRecordSet!amount; Tab(6); Str; Tab(22 + (8 - Len(MyRecordSet!FeeUnit)) / 2); MyRecordSet!FeeUnit; Tab(30 + (8 - Len(MyRecordSet!FeeTotal)) / 2); MyRecordSet!FeeTotal
                            MyRecordSet.MoveNext
                        Loop
                    End If
                    Set MyRecordSet = RunParametricStoredProcedure2Rec("Get_InvoiceInfo", Parameter)
                    Printer.Font.size = 10
                    Printer.Print "---------------------------------------------------------------"  'LF
                   
                   
                   If MyRecordSet!DiscountTotal <> 0 Or MyRecordSet!CarryFeeTotal <> 0 Or MyRecordSet!PackingTotal <> 0 Or MyRecordSet!ServiceTotal <> 0 Or MyRecordSet!TaxTotal <> 0 Or MyRecordSet!DutyTotal <> 0 Then
                       Printer.Font.size = 12
                       Printer.Print "      Ã„⁄ ò· :  " & Format(Totalprice, "#,##  —Ì«·")   ' MyRecordSet!Sumprice & "    —Ì«·   "
                       'Printer.Print "-----------------------------------------------------------"  'LF
                   End If
                    Printer.Font.size = 9
                    If MyRecordSet!DiscountTotal <> 0 Then
                       'Printer.Print MyRecordSet!DiscountTotal & tempstring.ReadLine
                       Printer.Print "  Œ›Ì› : "; Format(MyRecordSet!DiscountTotal, "#,##  —Ì«·")  '& MyRecordSet!DiscountTotal & "    —Ì«·   "
                    End If
                    If MyRecordSet!CarryFeeTotal <> 0 Then
                       'Printer.Print MyRecordSet!CarryFeeTotal & tempstring.ReadLine
                       Printer.Print " ò—«ÌÂ Õ„· : "; Format(MyRecordSet!CarryFeeTotal, "#,##  —Ì«·")  '& MyRecordSet!CarryFeeTotal & "    —Ì«·   "
                    End If
                    If MyRecordSet!PackingTotal <> 0 Then
                       'Printer.Print MyRecordSet!PackingTotal & tempstring.ReadLine
                       Printer.Print "  »” Â »‰œÌ : "; Format(MyRecordSet!PackingTotal, "#,##  —Ì«·") '& MyRecordSet!PackingTotal & "    —Ì«·   "
                   End If
                     If MyRecordSet!ServiceTotal <> 0 Then
                        'Printer.Print MyRecordSet!PackingTotal & tempstring.ReadLine
                       Printer.Print "  ”—ÊÌ” : " & Format(MyRecordSet!ServiceTotal, "#,##  —Ì«·")  'MyRecordSet!taxTotal & "    —Ì«·   "
                    End If
                   If MyRecordSet!TaxTotal <> 0 Then
                      'Printer.Print MyRecordSet!ServiceTotal & tempstring.ReadLine
                      Printer.Print "  „«·Ì«  : " & Format(MyRecordSet!TaxTotal, "#,##  —Ì«·")  'MyRecordSet!taxTotal & "    —Ì«·   "
                   End If
                   If MyRecordSet!DutyTotal <> 0 Then
                      'Printer.Print MyRecordSet!ServiceTotal & tempstring.ReadLine
                      Printer.Print "  ⁄Ê«—÷ : " & Format(MyRecordSet!DutyTotal, "#,##  —Ì«·")  '& MyRecordSet!DutyTotal & "    —Ì«·   "
                   End If
                   
                   If MyRecordSet!DiscountTotal <> 0 Or MyRecordSet!CarryFeeTotal <> 0 Or MyRecordSet!PackingTotal <> 0 Or MyRecordSet!ServiceTotal <> 0 Or MyRecordSet!TaxTotal <> 0 Or MyRecordSet!DutyTotal <> 0 Then
                       Printer.Font.size = 10
                       Printer.Print "-----------------------------------------------------------"  'LF
                   End If
                   Printer.FontName = "Titr"  ' FontA1x1   Trrafic
                   Printer.Font.size = 12
                   Printer.Font.Bold = True
                   'Printer.Print "    " & tempstring.ReadLine & ":  " & MyRecordSet!Sumprice & "    —Ì«·   "
                   Str = "      ﬁ«»· Å—œ«Œ  :  " & Format(MyRecordSet!sumPrice, "#,##  —Ì«·")  ' MyRecordSet!Sumprice & "    —Ì«·   "
                   Printer.Print Space((60 - Len(Str)) / 2) & Str
                   
                   Printer.Font.size = 10
                   Printer.Print "------------------------------------------------------------"  'LF
        
                   If Trim(MyRecordSet!NoticeDescription) <> "" Then
                       Printer.FontName = "Nazanin"  ' FontA1x1   Trrafic
                       Printer.Font.size = 12
                       Printer.Print Space((60 - Len(MyRecordSet!NoticeDescription)) / 2) & MyRecordSet!NoticeDescription
                   End If
                   
                   If Trim(MyRecordSet!TempAddress) <> "" Then
                       Printer.FontName = "Nazanin"  ' FontA1x1   Trrafic
                       Printer.Font.size = 12
                       Printer.Print Space((60 - Len(MyRecordSet!TempAddress)) / 2) & MyRecordSet!TempAddress
                       Printer.Print "                                     "
                   End If
                   Printer.EndDoc
                    
                    If InStr(1, PassedPrinterName, "Samsung", 1) Then

                        If MyRecordSet!barcode <> "" Then
                           Printer.Font.size = 20
                           Printer.FontName = "Code39"
                           Printer.Print MyRecordSet!barcode
                        End If
                         If MyRecordSet!LineFeed <> "" Then
                            Printer.Font.size = 9.5
                            Printer.FontName = "FontControl"
                           Printer.Print MyRecordSet!LineFeed
                        End If
                       If MyRecordSet!Cutter <> "" Then
                            Printer.Font.size = 9.5
                            Printer.FontName = "FontControl"
                           Printer.Print MyRecordSet!Cutter
                        End If
                        If MyRecordSet!Arm <> "" Then
                            Printer.Font.size = 9.5
                            Printer.FontName = "FontControl"
                           Printer.Print MyRecordSet!Arm
                        End If
                        Printer.Font.Bold = True
                        Printer.Font.size = 14
                        Printer.FontName = "Trrafic"  '
                       
                       Printer.EndDoc
                   
                   ElseIf InStr(1, PassedPrinterName, "Tsc", 1) Then
                        If MyRecordSet!barcode <> "" Then
                        End If
                         If MyRecordSet!LineFeed <> "" Then
                        End If
                       If MyRecordSet!Cutter <> "" Then
                        End If
                        If MyRecordSet!Arm <> "" Then
                        End If
                        Printer.Font.Bold = True
                        Printer.Font.size = 14
                        Printer.FontName = "Trrafic"  '
                   ElseIf InStr(1, PassedPrinterName, "Series 280", 1) Then

                        Open Printer.Port For Output As #1
                        
                        
                        If MyRecordSet!barcode <> "" Then
''''                           Printer.Font.Size = 20
''''                           Printer.FontName = "Code128"
''''                           Printer.Print MyRecordSet!barcode
                        End If
                         If MyRecordSet!LineFeed <> "" Then
                           Print #1, Chr$(27) & Chr$(100) & Chr$(Len(MyRecordSet!LineFeed))   ' Line Feed (5 Line)
                        End If
                       If MyRecordSet!Cutter <> "" Then
                           Print #1, Chr$(27) & Chr$(118)  ' Cut
                        End If
'                        If MyRecordSet!Arm <> "" Then
'                           Print #1, Chr$(27) & Chr$(31) & "r" & ARM1 & Chr$(0) ' Print Arm
'                        End If
                        
                        Close #1

                   End If
                    
            Case 2      'LowLevel Printer Output
                                
                Dim DoubleHigh, NormalHigh As String
                ReDim Parameter(9) As Parameter
                Parameter(0) = GenerateInputParameter("@intLanguage", adBoolean, 1, clsStation.Language)
                Parameter(1) = GenerateInputParameter("@intFacNO", adBigInt, 8, FichNumber)
                Parameter(2) = GenerateInputParameter("@PrintFormat", adInteger, 4, PrintFormat)
                Parameter(3) = GenerateInputParameter("@StationID", adInteger, 4, intStationId)
                Parameter(4) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                Parameter(5) = GenerateInputParameter("@intPrinterNo", adInteger, 4, PrinterNo)
                Parameter(6) = GenerateInputParameter("@Mode", adInteger, 4, AddEditMode)
                Parameter(7) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                Parameter(8) = GenerateInputParameter("@PartitionId", adInteger, 4, PartitionNo)
                Parameter(9) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                Set MyRecordSet = RunParametricStoredProcedure2Rec("Get_InvoiceInfo", Parameter, ClsCnn)
                If InStr(1, ReportFileName, "InvoiceFich_Send", 1) Then
                   TxtReportFile = App.Path & "\Reports" & RepVer & "\" & "InvoiceFich_Send" & ".txt"
                    IsFileExist = filetemp.FileExists(TxtReportFile)
                    
                    If IsFileExist = False Then
                        frmDisMsg.lblMessage = "›«Ì· „ ‰Ì »—«Ì «Ì‰ ›—„  ç«Å  ⁄—Ì› ‰‘œÂ «”  "
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Timer1.Interval = 2000
                        frmDisMsg.Show vbModal
                    Else
                        On Error GoTo ErrorHandler
                        If InStr(1, Printer.DeviceName, "\\", 1) Then
                        '   Open "\\demo1\SAMSUNG SRP350\Lpt1" For Output As #1
                           Open Printer.DeviceName & "\" & PortName For Output As #1
                        Else
                           Open Printer.Port For Output As #1
                        End If
                        
                        Print #1, Chr$(27) & Chr(40)    ' Initialize
                        
                        For i = 1 To intNumberOfCopies
                        If InStr(1, PassedPrinterName, "Samsung", 1) Then
                            NormalHigh = Chr$(27) & Chr$(33) & Chr(0)    'Normal High
                            DoubleHigh = Chr$(27) & Chr$(33) & Chr(24)    'Double High
                        ElseIf InStr(1, PassedPrinterName, "Series 280", 1) Then
                            NormalHigh = Chr$(27) & Chr$(87) & Chr(0)    'Normal High
                            DoubleHigh = Chr$(27) & Chr$(87) & Chr(2)    'Double High
                        ElseIf InStr(1, PassedPrinterName, "TSC", 1) Then
                            NormalHigh = ""    'Normal High
                            DoubleHigh = ""    'Double High
                        End If
                        
                        Set tempstring = filetemp.OpenTextFile(TxtReportFile, ForReading, False, TristateFalse)
                        Print #1, NormalHigh
                        Print #1, "     " & WinToIranSys(MyRecordSet!RecursievAlert) & "          " & WinToIranSys(MyRecordSet!ReportHeder)
                     ''   Print #1, Chr$(27) & Chr$(33) & Chr(24)   'Double High
                        Str = tempstring.ReadLine
                        If Str <> "" And left(Str, 1) <> "-" Then            ' For Hiden StationId  From txt File
                            Str = "                   " & intStationId & WinToIranSys(Str)
                            Print #1, Str & vbCrLf
                        End If
                        Str = WinToIranSys(tempstring.ReadLine)
                        Str = "    " & WinToIranSys(left(MyRecordSet!Date, 2)) & "/" & WinToIranSys(Mid(MyRecordSet!Date, 4, 2)) & "/" & WinToIranSys(Right(MyRecordSet!Date, 2)) & " :" & WinToIranSys(tempstring.ReadLine) & "        " & DoubleHigh & WinToIranSys(Right(CStr(FichNumber), 3)) & " :" & Str
                        Print #1, Str & vbCrLf
                        Print #1, NormalHigh
                        Print #1, "------------------------------------------"  'LF
                        
                        Str = tempstring.ReadLine
                        If MyRecordSet!MembershipId <> -1 Then          ' For Hiden ServePlace From txt File
                           Print #1, DoubleHigh
                           Print #1, "" & WinToIranSys(MyRecordSet!Family) & " " & WinToIranSys(MyRecordSet!MembershipId) & " :" & WinToIranSys(Str)
                           Print #1, NormalHigh
                           Print #1, "******************************************"  'LF
                           Print #1, DoubleHigh
                           If Len(MyRecordSet!CustomerAddress) > 40 Then
                               Dim ggg As Integer
                               ggg = InStr(35, MyRecordSet!CustomerAddress, " ", 1)
                               Print #1, WinToIranSys(left(MyRecordSet!CustomerAddress, ggg))
                               Print #1, WinToIranSys(Mid(MyRecordSet!CustomerAddress, ggg + 1))
                           Else
                               Print #1, WinToIranSys(MyRecordSet!CustomerAddress)
                           End If
                           If Len(MyRecordSet!Tel1) <> 0 Then
                               Print #1, "              " & WinToIranSys(left(MyRecordSet!Tel1, 8)) & WinToIranSys(" ·›‰ ")
                           ElseIf Len(MyRecordSet!Tel2) <> 0 Then
                               Print #1, "              " & WinToIranSys(left(MyRecordSet!Tel1, 8)) & WinToIranSys(" ·›‰ ")
                           End If
                           Print #1, NormalHigh
                           Print #1, "******************************************"  'LF
                        End If
                        
                        If Not (MyRecordSet.BOF Or MyRecordSet.EOF) Then
                             Print #1, " " & WinToIranSys(tempstring.ReadLine) & " "
                             Print #1, "|----------------------------------------|"  'LF
     '                        Printer.Print " „ﬁœ«—            ‰«„ ﬂ«·«            ›Ì            Ã„⁄"
                             Do While Not (MyRecordSet.EOF)
                               '  str = Right(MyRecordSet!AsciiNamePrn, 15)
                                 Str = Right(WinToIranSys(MyRecordSet!NamePrn), 20)
                                 Print #1, "|" & Space(7 - Len(MyRecordSet!FeeTotal)) & WinToIranSys(MyRecordSet!FeeTotal) & Space(7 - Len(MyRecordSet!FeeUnit)) & WinToIranSys(MyRecordSet!FeeUnit) & Space(20 - Len(Trim(Str))) & Str & Space(6 - Len(MyRecordSet!amount)) & WinToIranSys(MyRecordSet!amount) & "|"
                                 
                                 MyRecordSet.MoveNext
                                 
                             Loop
                         End If
                         Print #1, "------------------------------------------"  'LF
                        
                         Set MyRecordSet = RunParametricStoredProcedure2Rec("Get_InvoiceInfo", Parameter, ClsCnn)
                         If MyRecordSet!DiscountTotal < 0 Then
                            Print #1, "  " & WinToIranSys(MyRecordSet!DiscountTotal) & WinToIranSys(tempstring.ReadLine)
                         Else
                            Str = tempstring.ReadLine
                         End If
                         If MyRecordSet!CarryFeeTotal <> 0 Then
                            Print #1, "  " & WinToIranSys(MyRecordSet!CarryFeeTotal) & WinToIranSys(tempstring.ReadLine)
                         Else
                            Str = tempstring.ReadLine
                         End If
                         If MyRecordSet!PackingTotal <> 0 Then
                            Print #1, "  " & WinToIranSys(MyRecordSet!PackingTotal) & WinToIranSys(tempstring.ReadLine)
                         Else
                           Str = tempstring.ReadLine
                        End If
                        If MyRecordSet!ServiceTotal <> 0 Then
                           Print #1, "  " & WinToIranSys(MyRecordSet!ServiceTotal) & WinToIranSys(tempstring.ReadLine)
                        Else
                           Str = tempstring.ReadLine
                        End If
                        
                        If MyRecordSet!DiscountTotal <> 0 Or MyRecordSet!CarryFeeTotal <> 0 Or MyRecordSet!PackingTotal <> 0 Or MyRecordSet!ServiceTotal <> 0 Then
                            Print #1, "------------------------------------------"  'LF
                        End If
                        Print #1, DoubleHigh & "     " & WinToIranSys("   —Ì«·   ") & WinToIranSys(MyRecordSet!sumPrice) & "   :   " & WinToIranSys(tempstring.ReadLine)
                        
                        Print #1, NormalHigh & "------------------------------------------"  'LF
    
                        If MyRecordSet!NoticeDescription <> "" Then
                           If Len(MyRecordSet!NoticeDescription) < 40 Then
                               Print #1, Space(20 - Len(MyRecordSet!NoticeDescription) / 2) & WinToIranSys(MyRecordSet!NoticeDescription)
                           Else
                               ggg = InStr(35, MyRecordSet!NoticeDescription, " ", 1)
                               Print #1, WinToIranSys(left(MyRecordSet!NoticeDescription, ggg))
                               Print #1, WinToIranSys(Mid(MyRecordSet!NoticeDescription, ggg + 1))
                           End If
                        End If
                        Print #1, "              "
                        
                        If InStr(1, PassedPrinterName, "Samsung", 1) Then
                        '   Print #1, Chr$(30)    ' Beep
                        '   Print #1, Chr$(27) & Chr(40)    ' Initialize
                             If MyRecordSet!barcode <> "" Then
                                Print #1, Chr$(29) & Chr$(104) & Chr$(80)    'Height
                                Print #1, Chr$(29) & Chr$(119) & Chr(3) 'Width
                                Print #1, Chr$(29) & Chr$(107) & Chr$(2) & MyRecordSet!barcode & Chr$(0)
                                Print #1, "     " & Right(MyRecordSet!barcode, 10)
                             End If
                             If MyRecordSet!LineFeed <> "" Then
                                Print #1, Chr$(27) & Chr$(100) & Chr$(Len(MyRecordSet!LineFeed))
                             End If
                             If MyRecordSet!Cutter <> "" Then
                                Print #1, Chr$(27) & Chr$(109)  ' Cut
                             End If
                             If MyRecordSet!Arm <> "" Then
                                Print #1, Chr$(28) & Chr$(112) & Chr$(1) & Chr$(0) 'Print Arm
                             End If
                        
                       ElseIf InStr(1, PassedPrinterName, "Tsc", 1) Then
                            If MyRecordSet!barcode <> "" Then
                            End If
                             If MyRecordSet!LineFeed <> "" Then
                            End If
                           If MyRecordSet!Cutter <> "" Then
                            End If
                            If MyRecordSet!Arm <> "" Then
                            End If
                       ElseIf InStr(1, PassedPrinterName, "Series 280", 1) Then
    
                             If MyRecordSet!barcode <> "" Then
''''                                 Print #1, Chr$(27) & Chr$(25) & Chr$(66), Chr$(1)    'Height
''''                                 Print #1, Chr$(27) & Chr$(25) & Chr(87), Chr$(3)  'Width
                                 Print #1, Chr$(27) & Chr$(98) & Chr$(4) & MyRecordSet!barcode & Chr$(3)
                                 Print #1, "             " & Right(MyRecordSet!barcode, 10)
                             End If
                              If MyRecordSet!LineFeed <> "" Then
                                Print #1, Chr$(27) & Chr$(100) & Chr$(Len(MyRecordSet!LineFeed))   ' Line Feed (5 Line)
                             End If
                            If MyRecordSet!Cutter <> "" Then
                                Print #1, Chr$(27) & Chr$(118)  ' Cut
                             End If
'                             If MyRecordSet!Arm <> "" Then
'                                Print #1, Chr$(27) & Chr$(31) & "r" & ARM1 & Chr$(0) ' Print Arm
'                             End If
                            
    
                       End If
                       tempstring.Close
                       
                       Next i
                       Close #1
          
                   End If
                ElseIf InStr(1, ReportFileName, "InvoiceFich", 1) Then    '
'                TxtReportFile = Left(ReportFileName, Len(ReportFileName) - 4)
                    TxtReportFile = App.Path & "\Reports" & RepVer & "\" & "InvoiceFich" & ".txt"
                    
                    IsFileExist = filetemp.FileExists(TxtReportFile)
                    
                    If IsFileExist = False Then
                        frmDisMsg.lblMessage = "›«Ì· „ ‰Ì »—«Ì «Ì‰ ›—„  ç«Å  ⁄—Ì› ‰‘œÂ «”  "
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Timer1.Interval = 2000
                        frmDisMsg.Show vbModal
                    Else
                        On Error GoTo ErrorHandler
                        If InStr(1, Printer.DeviceName, "\\", 1) Then
                        '   Open "\\demo1\SAMSUNG SRP350\Lpt1" For Output As #1
                           Open Printer.DeviceName & "\" & PortName For Output As #1
                        Else
                           Open Printer.Port For Output As #1
                        End If
                        
                        Print #1, Chr$(27) & Chr(40)    ' Initialize
                        
                        For i = 1 To intNumberOfCopies
                        
                        If InStr(1, PassedPrinterName, "Samsung", 1) Then
                            NormalHigh = Chr$(27) & Chr$(33) & Chr(0)    'Normal High
                            DoubleHigh = Chr$(27) & Chr$(33) & Chr(24)    'Double High
                        ElseIf InStr(1, PassedPrinterName, "Series 280", 1) Then
                            NormalHigh = Chr$(27) & Chr$(87) & Chr(0)    'Normal High
                            DoubleHigh = Chr$(27) & Chr$(87) & Chr(2)    'Double High
                        ElseIf InStr(1, PassedPrinterName, "TSC", 1) Then
                            NormalHigh = ""    'Normal High
                            DoubleHigh = ""    'Double High
                        End If
                        
                        Set tempstring = filetemp.OpenTextFile(TxtReportFile, ForReading, False, TristateFalse)
                        Print #1, NormalHigh
                     ''   Print #1, Chr$(27) & Chr$(33) & Chr(24)   'Double High
                        If IsNull(MyRecordSet!RecursievAlert) = False And IsNull(MyRecordSet!ReportHeder) = False Then
                            Print #1, "     " & WinToIranSys(MyRecordSet!RecursievAlert) & "          " & WinToIranSys(MyRecordSet!ReportHeder)
                        End If
                        Str = tempstring.ReadLine
                        If Str <> "" And left(Str, 1) <> "-" Then            ' For Hiden StationId  From txt File
                            Str = "                   " & intStationId & WinToIranSys(Str)
                            Print #1, Str & vbCrLf
                        End If
                        Str = WinToIranSys(tempstring.ReadLine)
                        'str = "    " & WinToIranSys(Left(MyRecordSet!Date, 2)) & "/" & WinToIranSys(Mid(MyRecordSet!Date, 4, 2)) & "/" & WinToIranSys(Right(MyRecordSet!Date, 2)) & " :" & WinToIranSys(tempstring.ReadLine) & "          " & DoubleHigh & WinToIranSys(Right(CStr(FichNumber), 3)) & " :" & str
                        Str = "    " & left(MyRecordSet!Date, 2) & "/" & Mid(MyRecordSet!Date, 4, 2) & "/" & Right(MyRecordSet!Date, 2) & " :" & WinToIranSys(tempstring.ReadLine) & "          " & DoubleHigh & Right(CStr(FichNumber), 3) & " :" & Str
                        Print #1, Str & vbCrLf
                        Print #1, NormalHigh
                        Print #1, "------------------------------------------"  'LF
                        
                        Str = tempstring.ReadLine
                        If Str <> "" And left(Str, 1) <> "-" Then          ' For Hiden ServePlace From txt File
                           Print #1, "                " & WinToIranSys(MyRecordSet!ItemServeplaceDesc) & "  : " & WinToIranSys(Str)
                           Print #1, "------------------------------------------"  'LF
                        End If
                        
                        If Not (MyRecordSet.BOF Or MyRecordSet.EOF) Then
                             Print #1, "|" & WinToIranSys(tempstring.ReadLine) & "|"
                             Print #1, "|----------------------------------------|"  'LF
     '                        Printer.Print " „ﬁœ«—            ‰«„ ﬂ«·«            ›Ì            Ã„⁄"
                             Do While Not (MyRecordSet.EOF)
                               '  str = Right(MyRecordSet!AsciiNamePrn, 15)
                                 Str = Right(WinToIranSys(MyRecordSet!NamePrn), 20)
                                 Print #1, "|" & Space(7 - Len(MyRecordSet!FeeTotal)) & WinToIranSys(MyRecordSet!FeeTotal) & Space(7 - Len(MyRecordSet!FeeUnit)) & WinToIranSys(MyRecordSet!FeeUnit) & Space(20 - Len(Trim(Str))) & Str & Space(6 - Len(MyRecordSet!amount)) & WinToIranSys(MyRecordSet!amount) & "|"
                                 
                                 MyRecordSet.MoveNext
                                 
                             Loop
                         End If
                         Print #1, "------------------------------------------"  'LF
                        
                         Set MyRecordSet = RunParametricStoredProcedure2Rec("Get_InvoiceInfo", Parameter, ClsCnn)
                         If MyRecordSet!DiscountTotal <> 0 Then
                            Print #1, "  " & WinToIranSys(MyRecordSet!DiscountTotal) & WinToIranSys(tempstring.ReadLine)
                         Else
                            Str = tempstring.ReadLine
                         End If
                         If MyRecordSet!CarryFeeTotal <> 0 Then
                            Print #1, "  " & WinToIranSys(MyRecordSet!CarryFeeTotal) & WinToIranSys(tempstring.ReadLine)
                         Else
                            Str = tempstring.ReadLine
                         End If
                         If MyRecordSet!PackingTotal <> 0 Then
                            Print #1, "  " & WinToIranSys(MyRecordSet!PackingTotal) & WinToIranSys(tempstring.ReadLine)
                         Else
                           Str = tempstring.ReadLine
                        End If
                        If MyRecordSet!ServiceTotal <> 0 Then
                           Print #1, "  " & WinToIranSys(MyRecordSet!ServiceTotal) & WinToIranSys(tempstring.ReadLine)
                        Else
                           Str = tempstring.ReadLine
                        End If
                        
                        If MyRecordSet!DiscountTotal <> 0 Or MyRecordSet!CarryFeeTotal <> 0 Or MyRecordSet!PackingTotal <> 0 Or MyRecordSet!ServiceTotal <> 0 Then
                            Print #1, "------------------------------------------"  'LF
                        End If
                        Print #1, DoubleHigh & "     " & WinToIranSys("   —Ì«·   ") & WinToIranSys(MyRecordSet!sumPrice) & "   :   " & WinToIranSys(tempstring.ReadLine)
                        
                        Print #1, NormalHigh & "------------------------------------------"   'LF
    
                        If MyRecordSet!NoticeDescription <> "" Then
                           If Len(MyRecordSet!NoticeDescription) < 40 Then
                               Print #1, Space(20 - Len(MyRecordSet!NoticeDescription) / 2) & WinToIranSys(MyRecordSet!NoticeDescription)
                           Else
                               ggg = InStr(35, MyRecordSet!NoticeDescription, " ", 1)
                               Print #1, WinToIranSys(left(MyRecordSet!NoticeDescription, ggg))
                               Print #1, WinToIranSys(Mid(MyRecordSet!NoticeDescription, ggg + 1))
                           End If
                        End If
                        Print #1, "              "
                        
                        If InStr(1, PassedPrinterName, "Samsung", 1) Then
                        '    Print #1, Chr$(30)    ' Beep
                        '   Print #1, Chr$(27) & Chr(40)    ' Initialize
                             If MyRecordSet!barcode <> "" Then
                                Print #1, Chr$(29) & Chr$(104) & Chr$(80)    'Height
                                Print #1, Chr$(29) & Chr$(119) & Chr(3) 'Width
                                Print #1, "      " & Right(MyRecordSet!barcode, 10)
                                Print #1, Chr$(29) & Chr$(107) & Chr$(2) & MyRecordSet!barcode & Chr$(0)
                             End If
                             If MyRecordSet!LineFeed <> "" Then
                                Print #1, Chr$(27) & Chr$(100) & Chr$(Len(MyRecordSet!LineFeed))
                             End If
                             If MyRecordSet!Cutter <> "" Then
                                Print #1, Chr$(27) & Chr$(109)  ' Cut
                             End If
                             If MyRecordSet!Arm <> "" Then
                                Print #1, Chr$(28) & Chr$(112) & Chr$(1) & Chr$(0) 'Print Arm
                             End If
                        
                       ElseIf InStr(1, PassedPrinterName, "Tsc", 1) Then
                            If MyRecordSet!barcode <> "" Then
                            End If
                             If MyRecordSet!LineFeed <> "" Then
                            End If
                           If MyRecordSet!Cutter <> "" Then
                            End If
                            If MyRecordSet!Arm <> "" Then
                            End If
                       ElseIf InStr(1, PassedPrinterName, "Series 280", 1) Then
    
                         If MyRecordSet!barcode <> "" Then
''''                            Print #1, Chr$(27) & Chr$(25) & Chr$(66), Chr$(4)    'Height
''''                            Print #1, Chr$(27) & Chr$(25) & Chr(87), Chr$(4)  'Width
                            Print #1, Chr$(27) & Chr$(98) & Chr$(4) & MyRecordSet!barcode & Chr$(3)
                            Print #1, "      " & Right(MyRecordSet!barcode, 10)
                         End If
                          If MyRecordSet!LineFeed <> "" Then
                            Print #1, Chr$(27) & Chr$(100) & Chr$(Len(MyRecordSet!LineFeed))   ' Line Feed (5 Line)
                         End If
                        If MyRecordSet!Cutter <> "" Then
                            Print #1, Chr$(27) & Chr$(118)  ' Cut
                         End If
'                         If MyRecordSet!Arm <> "" Then
'                            Print #1, Chr$(27) & Chr$(31) & "r" & ARM1 & Chr$(0) ' Print Arm
'                         End If
                            
    
                       End If
                       tempstring.Close
                       
                       Next i
                       Close #1
          
                   End If
                
                ElseIf InStr(1, ReportFileName, "KitchenFich", 1) Then  ' '›Ì‘ ¬‘Å“Œ«‰Â
                    
                    TxtReportFile = App.Path & "\Reports" & RepVer & "\" & "KitchenFich" & ".txt"
                    
                    IsFileExist = filetemp.FileExists(TxtReportFile)
                    
                    If IsFileExist = False Then
                        frmDisMsg.lblMessage = "›«Ì· „ ‰Ì »—«Ì «Ì‰ ›—„  ç«Å  ⁄—Ì› ‰‘œÂ «”  "
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Timer1.Interval = 2000
                        frmDisMsg.Show vbModal
                    Else
                        On Error GoTo ErrorHandler
                        If InStr(1, Printer.DeviceName, "\\", 1) Then
                        '   Open "\\demo1\SAMSUNG SRP350\Lpt1" For Output As #1
                           Open Printer.DeviceName & "\" & PortName For Output As #1
                        Else
                           Open Printer.Port For Output As #1
                        End If
                        
                        For i = 1 To intNumberOfCopies
                        
                        If InStr(1, PassedPrinterName, "Samsung", 1) Then
                            NormalHigh = Chr$(27) & Chr$(33) & Chr(0)    'Normal High
                            DoubleHigh = Chr$(27) & Chr$(33) & Chr(24)    'Double High
                        ElseIf InStr(1, PassedPrinterName, "Series 280", 1) Then
                            NormalHigh = Chr$(27) & Chr$(87) & Chr(0)    'Normal High
                            DoubleHigh = Chr$(27) & Chr$(87) & Chr(2)    'Double High
                        ElseIf InStr(1, PassedPrinterName, "TSC", 1) Then
                            NormalHigh = ""    'Normal High
                            DoubleHigh = ""    'Double High
                        End If
                        
                        Set tempstring = filetemp.OpenTextFile(TxtReportFile, ForReading, False, TristateFalse)
                      '  Print #1, Chr$(27) & Chr$(71) & Chr(1) 'Double Strike
                       ' Print #1, Chr$(27) & "I" & Chr(1) '
                        Print #1, DoubleHigh
                      ' Print #1, Chr$(27) & Chr$(77) & Chr(0) 'Font A
                        Print #1, "     " & WinToIranSys(MyRecordSet!RecursievAlert) & "          " & WinToIranSys(MyRecordSet!ReportHeder)
                        Str = tempstring.ReadLine
                        If Str <> "" And left(Str, 1) <> "-" Then            ' For Hiden StationId  From txt File
                            Str = "                   " & intStationId & WinToIranSys(Str)
                            Print #1, Str & vbCrLf
                        End If
                        Str = WinToIranSys(tempstring.ReadLine)
                        Str = "    " & WinToIranSys(left(MyRecordSet!Date, 2)) & "/" & WinToIranSys(Mid(MyRecordSet!Date, 4, 2)) & "/" & WinToIranSys(Right(MyRecordSet!Date, 2)) & "  :" & WinToIranSys(tempstring.ReadLine) & "          " & WinToIranSys(Right(CStr(FichNumber), 3)) & "  :" & Str
                        Print #1, Str & vbCrLf
                        Print #1, "------------------------------------------"  'LF
                        
                        Str = tempstring.ReadLine
                        If Str <> "" And left(Str, 1) <> "-" Then            ' For Hiden ServePlace From txt File
                           Print #1, "                   " & WinToIranSys(MyRecordSet!FactorServeDescription) & "  : " & WinToIranSys(Str)
                           Print #1, "------------------------------------------"  'LF
                        End If
                        Str = tempstring.ReadLine
                        
                        If Not (MyRecordSet.BOF Or MyRecordSet.EOF) Then
                             Print #1, WinToIranSys(tempstring.ReadLine)
                             Print #1, "------------------------------------------"  'LF
                             Do While Not (MyRecordSet.EOF)
                               '  str = Right(MyRecordSet!AsciiNamePrn, 15)
                                 Str = Right(WinToIranSys(MyRecordSet!GoodName), 20)
                                 Print #1, Space(12 - Len(MyRecordSet!DifferencesDescription)) & WinToIranSys(MyRecordSet!DifferencesDescription) & Space(20 - Len(Trim(Str))) & Str & Space(7 - Len(MyRecordSet!amount)) & WinToIranSys(MyRecordSet!amount)
                                 
                                 MyRecordSet.MoveNext
                                 
                             Loop
                         End If
                         Set MyRecordSet = RunParametricStoredProcedure2Rec("Get_InvoiceInfo", Parameter, ClsCnn)
                         Print #1, "------------------------------------------"  'LF
                        
    
                        If MyRecordSet!NoticeDescription <> "" Then
                           If Len(MyRecordSet!NoticeDescription) < 40 Then
                               Print #1, Space(20 - Len(MyRecordSet!NoticeDescription) / 2) & WinToIranSys(MyRecordSet!NoticeDescription)
                           Else
                               ggg = InStr(35, MyRecordSet!NoticeDescription, " ", 1)
                               Print #1, WinToIranSys(left(MyRecordSet!NoticeDescription, ggg))
                               Print #1, WinToIranSys(Mid(MyRecordSet!NoticeDescription, ggg + 1))
                           End If
                        End If
                        Print #1, "              "
                        
                        If InStr(1, PassedPrinterName, "Samsung", 1) Then
                        '    Print #1, Chr$(30)    ' Beep
                        '   Print #1, Chr$(27) & Chr(40)    ' Initialize
                             If MyRecordSet!barcode <> "" Then
                                Print #1, Chr$(29) & Chr$(104) & Chr$(80)    'Height
                                Print #1, Chr$(29) & Chr$(119) & Chr(3) 'Width
                                Print #1, Chr$(29) & Chr$(107) & Chr$(2) & MyRecordSet!barcode & Chr$(0)
                                Print #1, "      " & Right(MyRecordSet!barcode, 10)
                             End If
                             If MyRecordSet!LineFeed <> "" Then
                                Print #1, Chr$(27) & Chr$(100) & Chr$(Len(MyRecordSet!LineFeed))
                             End If
                             If MyRecordSet!Cutter <> "" Then
                                Print #1, Chr$(27) & Chr$(109)  ' Cut
                             End If
                             If MyRecordSet!Arm <> "" Then
                                Print #1, Chr$(28) & Chr$(112) & Chr$(1) & Chr$(0) 'Print Arm
                             End If
                        
                       ElseIf InStr(1, PassedPrinterName, "Tsc", 1) Then
                            If MyRecordSet!barcode <> "" Then
                            End If
                             If MyRecordSet!LineFeed <> "" Then
                            End If
                           If MyRecordSet!Cutter <> "" Then
                            End If
                            If MyRecordSet!Arm <> "" Then
                            End If
                            Printer.Font.Bold = True
                            Printer.Font.size = 14
                            Printer.FontName = "Trrafic"  '
                       ElseIf InStr(1, PassedPrinterName, "Series 280", 1) Then
    
                            If MyRecordSet!barcode <> "" Then
''''                                 Print #1, Chr$(27) & Chr$(25) & Chr$(66), Chr$(4)    'Height
''''                                 Print #1, Chr$(27) & Chr$(25) & Chr(87), Chr$(4)  'Width
                                 Print #1, Chr$(27) & Chr$(98) & Chr$(4) & MyRecordSet!barcode & Chr$(3)
                                 Print #1, "     " & Right(MyRecordSet!barcode, 10)
                            End If
                             If MyRecordSet!LineFeed <> "" Then
                               Print #1, Chr$(27) & Chr$(100) & Chr$(Len(MyRecordSet!LineFeed))   ' Line Feed (5 Line)
                            End If
                           If MyRecordSet!Cutter <> "" Then
                               Print #1, Chr$(27) & Chr$(118)  ' Cut
                            End If
'                            If MyRecordSet!Arm <> "" Then
'                               Print #1, Chr$(27) & Chr$(31) & "r" & ARM1 & Chr$(0) ' Print Arm
'                            End If
                            
    
                       End If
                       tempstring.Close
                       
                       Next i
                       Close #1
          
               
                   End If
             
             End If
         End Select
        
    End If
    
    Set MyRecordSet = Nothing
Exit Function

ErrorHandler:
    MsgBox err.Description & " - ç«Å ›Ì‘"
    Close #1

End Function

Public Function InsertPrintFich(FichNumber As Long, PrintFormat As Long, intStationId As Integer, PrinterNo, AddEditMode As EnumAddEditMode, PartitionNo) As Boolean
    
    Dim intPrintFichNo As Long
    On Error GoTo ErrorHandler
    ReDim Parameter(10) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 8, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@intFacNo", adInteger, 4, FichNumber)
    Parameter(2) = GenerateInputParameter("@PrintFormat", adInteger, 4, PrintFormat)
    Parameter(3) = GenerateInputParameter("@StationId", adInteger, 4, intStationId)
    Parameter(4) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(5) = GenerateInputParameter("@intPrinterNo", adInteger, 4, PrinterNo)
    Parameter(6) = GenerateInputParameter("@Mode", adInteger, 4, AddEditMode)
    Parameter(7) = GenerateInputParameter("@AccountYear", adInteger, 4, AccountYear)
    Parameter(8) = GenerateInputParameter("@PartitionId", adInteger, 4, PartitionNo)
    Parameter(9) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(10) = GenerateOutputParameter("@intPrintFichNo", adBigInt, 8)
    
    intPrintFichNo = RunParametricStoredProcedure("Insert_tblTotal_PrintFich", Parameter, ClsCnn)
    ''InsertPrintFich = intPrintFichNo
    InsertPrintFich = True
    Exit Function

ErrorHandler:
    InsertPrintFich = False
    InsertPrintFich = -1
    MsgBox err.Description
    
End Function

Private Sub Class_Initialize()
    ClsCnn.ConnectionString = strConnectionString
    ClsCnn.Open
End Sub

Private Sub Class_Terminate()
    If ClsCnn.State = adStateOpen Then ClsCnn.Close: Set ClsCnn = Nothing
    Set fs = Nothing
End Sub




