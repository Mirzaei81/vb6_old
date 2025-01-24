Attribute VB_Name = "ModSpool"
Option Explicit
Public PrinterNo(1 To 6) As Integer
Public PrinterName(1 To 6) As String

Public Declare Function lstrcpy Lib "kernel32" _
   Alias "lstrcpyA" _
   (ByVal lpString1 As String, _
   ByVal lpString2 As String) _
   As Long

Public Declare Function OpenPrinter Lib "winspool.drv" _
   Alias "OpenPrinterA" _
   (ByVal pPrinterName As String, _
   phPrinter As Long, _
   pDefault As PRINTER_DEFAULTS) _
   As Long

Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" _
   (ByVal hPrinter As Long, _
   ByVal Level As Long, _
   pPrinter As Byte, _
   ByVal cbBuf As Long, _
   pcbNeeded As Long) _
   As Long

Public Declare Function ClosePrinter Lib "winspool.drv" _
   (ByVal hPrinter As Long) _
   As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

Public Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" _
   (ByVal hPrinter As Long, _
   ByVal FirstJob As Long, _
   ByVal NoJobs As Long, _
   ByVal Level As Long, _
   pJob As Byte, _
   ByVal cdBuf As Long, _
   pcbNeeded As Long, _
   pcReturned As Long) _
   As Long
   
' constants for PRINTER_DEFAULTS structure
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ACCESS_ADMINISTER = &H4

' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

Public Type PRINTER_DEFAULTS
   pDatatype As String
   pDevMode As Long
   DesiredAccess As Long
End Type

Public Type DEVMODE
   dmDeviceName As String * CCHDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCHFORMNAME
   dmLogPixels As Integer
   dmBitsPerPel As Long
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Type JOB_INFO_2
   JobId As Long
   pPrinterName As Long
   pMachineName As Long
   pUserName As Long
   pDocument As Long
   pNotifyName As Long
   pDatatype As Long
   pPrintProcessor As Long
   pParameters As Long
   pDriverName As Long
   pDevMode As Long
   pStatus As Long
   pSecurityDescriptor As Long
   Status As Long
   Priority As Long
   Position As Long
   StartTime As Long
   UntilTime As Long
   TotalPages As Long
   Size As Long
   Submitted As SYSTEMTIME
   time As Long
   PagesPrinted As Long
End Type

Type PRINTER_INFO_2
   pServerName As Long
   pPrinterName As Long
   pShareName As Long
   pPortName As Long
   pDriverName As Long
   pComment As Long
   pLocation As Long
   pDevMode As Long
   pSepFile As Long
   pPrintProcessor As Long
   pDatatype As Long
   pParameters As Long
   pSecurityDescriptor As Long
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   cJobs As Long
   AveragePPM As Long
End Type

Public Const ERROR_INSUFFICIENT_BUFFER = 122
Public Const PRINTER_STATUS_BUSY = &H200
Public Const PRINTER_STATUS_DOOR_OPEN = &H400000
Public Const PRINTER_STATUS_ERROR = &H2
Public Const PRINTER_STATUS_INITIALIZING = &H8000
Public Const PRINTER_STATUS_IO_ACTIVE = &H100
Public Const PRINTER_STATUS_MANUAL_FEED = &H20
Public Const PRINTER_STATUS_NO_TONER = &H40000
Public Const PRINTER_STATUS_NOT_AVAILABLE = &H1000
Public Const PRINTER_STATUS_OFFLINE = &H80
Public Const PRINTER_STATUS_OUT_OF_MEMORY = &H200000
Public Const PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
Public Const PRINTER_STATUS_PAGE_PUNT = &H80000
Public Const PRINTER_STATUS_PAPER_JAM = &H8
Public Const PRINTER_STATUS_PAPER_OUT = &H10
Public Const PRINTER_STATUS_PAPER_PROBLEM = &H40
Public Const PRINTER_STATUS_PAUSED = &H1
Public Const PRINTER_STATUS_PENDING_DELETION = &H4
Public Const PRINTER_STATUS_PRINTING = &H400
Public Const PRINTER_STATUS_PROCESSING = &H4000
Public Const PRINTER_STATUS_TONER_LOW = &H20000
Public Const PRINTER_STATUS_USER_INTERVENTION = &H100000
Public Const PRINTER_STATUS_WAITING = &H2000
Public Const PRINTER_STATUS_WARMING_UP = &H10000
Public Const JOB_STATUS_PAUSED = &H1
Public Const JOB_STATUS_ERROR = &H2
Public Const JOB_STATUS_DELETING = &H4
Public Const JOB_STATUS_SPOOLING = &H8
Public Const JOB_STATUS_PRINTING = &H10
Public Const JOB_STATUS_OFFLINE = &H20
Public Const JOB_STATUS_PAPEROUT = &H40
Public Const JOB_STATUS_PRINTED = &H80
Public Const JOB_STATUS_DELETED = &H100
Public Const JOB_STATUS_BLOCKED_DEVQ = &H200
Public Const JOB_STATUS_USER_INTERVENTION = &H400
Public Const JOB_STATUS_RESTART = &H800

Private m_mainWmi As Object
Private m_deviceLists As Collection



Public Function GetString(ByVal PtrStr As Long) As String
   Dim StrBuff As String * 256
   
   'Check for zero address
   If PtrStr = 0 Then
      GetString = " "
      Exit Function
   End If
   
   'Copy data from PtrStr to buffer.
   CopyMemory ByVal StrBuff, ByVal PtrStr, 256
   
   'Strip any trailing nulls from string.
   GetString = StripNulls(StrBuff)
End Function

Public Function StripNulls(OriginalStr As String) As String
   'Strip any trailing nulls from input string.
   If (InStr(OriginalStr, Chr(0)) > 0) Then
      OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
   End If

   'Return modified string.
   StripNulls = OriginalStr
End Function

Public Function PtrCtoVbString(Add As Long) As String
    Dim sTemp As String * 512
    Dim X As Long

    X = lstrcpy(sTemp, Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
         PtrCtoVbString = ""
    Else
         PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function

Public Function CheckPrinterStatus(PI2Status As Long) As String
   Dim TempStr As String
   
   If PI2Status = 0 Then   ' Return "Ready"
      CheckPrinterStatus = "Printer Status = Ready" & vbCrLf
   Else
      TempStr = ""   ' Clear
      If (PI2Status And PRINTER_STATUS_BUSY) Then
         TempStr = TempStr & "Busy  "
      End If
      
      If (PI2Status And PRINTER_STATUS_DOOR_OPEN) Then
         TempStr = TempStr & "Printer Door Open  "
      End If
      
      If (PI2Status And PRINTER_STATUS_ERROR) Then
         TempStr = TempStr & "Printer Error  "
      End If
      
      If (PI2Status And PRINTER_STATUS_INITIALIZING) Then
         TempStr = TempStr & "Initializing  "
      End If
      
      If (PI2Status And PRINTER_STATUS_IO_ACTIVE) Then
         TempStr = TempStr & "I/O Active  "
      End If

      If (PI2Status And PRINTER_STATUS_MANUAL_FEED) Then
         TempStr = TempStr & "Manual Feed  "
      End If
      
      If (PI2Status And PRINTER_STATUS_NO_TONER) Then
         TempStr = TempStr & "No Toner  "
      End If
      
      If (PI2Status And PRINTER_STATUS_NOT_AVAILABLE) Then
         TempStr = TempStr & "Not Available  "
      End If
      
      If (PI2Status And PRINTER_STATUS_OFFLINE) Then
         TempStr = TempStr & "Off Line  "
      End If
      
      If (PI2Status And PRINTER_STATUS_OUT_OF_MEMORY) Then
         TempStr = TempStr & "Out of Memory  "
      End If
      
      If (PI2Status And PRINTER_STATUS_OUTPUT_BIN_FULL) Then
         TempStr = TempStr & "Output Bin Full  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PAGE_PUNT) Then
         TempStr = TempStr & "Page Punt  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PAPER_JAM) Then
         TempStr = TempStr & "Paper Jam  "
      End If

      If (PI2Status And PRINTER_STATUS_PAPER_OUT) Then
         TempStr = TempStr & "Paper Out  "
      End If
      
      If (PI2Status And PRINTER_STATUS_OUTPUT_BIN_FULL) Then
         TempStr = TempStr & "Output Bin Full  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PAPER_PROBLEM) Then
         TempStr = TempStr & "Page Problem  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PAUSED) Then
         TempStr = TempStr & "Paused  "
      End If

      If (PI2Status And PRINTER_STATUS_PENDING_DELETION) Then
         TempStr = TempStr & "Pending Deletion  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PRINTING) Then
         TempStr = TempStr & "Printing  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PROCESSING) Then
         TempStr = TempStr & "Processing  "
      End If
      
      If (PI2Status And PRINTER_STATUS_TONER_LOW) Then
         TempStr = TempStr & "Toner Low  "
      End If

      If (PI2Status And PRINTER_STATUS_USER_INTERVENTION) Then
         TempStr = TempStr & "User Intervention  "
      End If
      
      If (PI2Status And PRINTER_STATUS_WAITING) Then
         TempStr = TempStr & "Waiting  "
      End If
      
      If (PI2Status And PRINTER_STATUS_WARMING_UP) Then
         TempStr = TempStr & "Warming Up  "
      End If
      
      'Did you find a known status?
      If Len(TempStr) = 0 Then
         TempStr = "Unknown Status of " & PI2Status
      End If
      
      'Return the Status
      CheckPrinterStatus = "Printer Status = " & TempStr & vbCrLf
   End If
End Function

Private Function GetMainWMIObject() As Object
On Error GoTo EH
If m_mainWmi Is Nothing Then
Set m_mainWmi = GetObject("WinMgmts:")
End If
Set GetMainWMIObject = m_mainWmi
Exit Function
EH:
Set GetMainWMIObject = Nothing
End Function

Public Function WmiIsAvailable() As Boolean
WmiIsAvailable = CBool(Not GetMainWMIObject Is Nothing)
End Function

Public Function GetWmiDeviceSingleValue(ByVal WmiClass As String, ByVal WmiProperty As String) As String
On Error GoTo done
Dim Result As String

Dim wmiclassObjList As Object
Set wmiclassObjList = GetWmiDeviceList(WmiClass)
Dim wmiclassObj As Object
For Each wmiclassObj In wmiclassObjList
Result = CallByName(wmiclassObj, WmiProperty, VbGet)
Exit For
Next

done:
GetWmiDeviceSingleValue = Trim(Result)
End Function

Public Function GetWmiDeviceList(ByVal WmiClass As String) As Object
If m_deviceLists Is Nothing Then
Set m_deviceLists = New Collection
End If

On Error GoTo fetchNew

Set GetWmiDeviceList = m_deviceLists.Item(WmiClass)
Exit Function

fetchNew:
Dim devList As Object
Set devList = GetWmiDeviceListInternal(WmiClass)
If Not devList Is Nothing Then
Call m_deviceLists.Add(devList, WmiClass)
End If
Set GetWmiDeviceList = devList
End Function

Private Function GetWmiDeviceListInternal(ByVal WmiClass As String) As Object
    On Error GoTo EH
    Set GetWmiDeviceListInternal = GetMainWMIObject.Instancesof(WmiClass)
    Exit Function
EH:
    Set GetWmiDeviceListInternal = Nothing
End Function

Public Sub GetPrintersInDataBase()
        
    Dim i As Long
    Dim rctmp As ADODB.Recordset
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_tPrinters")
    
    For i = 1 To 6
        PrinterNo(i) = 0
        PrinterName(i) = ""
    Next
    i = 1
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        rctmp.MoveFirst
        Do While Not rctmp.EOF
            If i <= 6 Then
                PrinterNo(i) = rctmp!PrinterNo
                PrinterName(i) = CStr(rctmp!PrinterName)
                i = i + 1
                rctmp.MoveNext
            Else
                ShowDisMessage "ÝÞØ 6 ÑíäÊÑ ÞÇÏÑ Èå ãæäíÊæÑ ˜ÑÏä ãí ÈÇÔäÏ", 2000
                Exit Do
            End If
        Loop
    End If
    Set rctmp = Nothing

End Sub


Public Function CheckPrinter(PrinterName As String, PrinterStr As String, JobStr As String, JobQuantity As Long) As String
   Dim hPrinter As Long
   Dim ByteBuf As Long
   Dim BytesNeeded As Long
   Dim PI2 As PRINTER_INFO_2
   Dim JI2 As JOB_INFO_2
   Dim PrinterInfo() As Byte
   Dim JobInfo() As Byte
   Dim Result As Long
   Dim LastError As Long
   Dim TempStr As String
   Dim NumJI2 As Long
   Dim pDefaults As PRINTER_DEFAULTS
   Dim i As Integer
   
   'Set a default return value if no errors occur.
   CheckPrinter = "Printer info retrieved"
   
   'NOTE: You can pick a printer from the Printers Collection
   'or use the EnumPrinters() API to select a printer name.
   
   'Use the default printer of Printers collection.
   'This is typically, but not always, the system default printer.
   
   'Set desired access security setting.
   pDefaults.DesiredAccess = PRINTER_ACCESS_USE
   
   'Call API to get a handle to the printer.
   Result = OpenPrinter(PrinterName, hPrinter, pDefaults)
   If Result = 0 Then
      'If an error occurred, display an error and exit sub.
      CheckPrinter = "Cannot open printer " & PrinterName & _
         ", Error: " & err.LastDllError
      Exit Function
   End If

   'Init BytesNeeded
   BytesNeeded = 0

   'Clear the error object of any errors.
   err.Clear

   'Determine the buffer size that is needed to get printer info.
   Result = GetPrinter(hPrinter, 2, 0&, 0&, BytesNeeded)
   
   'Check for error calling GetPrinter.
   If err.LastDllError <> ERROR_INSUFFICIENT_BUFFER Then
      'Display an error message, close printer, and exit sub.
      CheckPrinter = " > GetPrinter Failed on initial call! <"
      ClosePrinter hPrinter
      Exit Function
   End If
   
   'Note that in Charles Petzold's book "Programming Windows 95," he
   'states that because of a problem with GetPrinter on Windows 95 only, you
   'must allocate a buffer as much as three times larger than the value
   'returned by the initial call to GetPrinter. This is not done here.
   ReDim PrinterInfo(1 To BytesNeeded)
   
   ByteBuf = BytesNeeded
   
   'Call GetPrinter to get the status.
   Result = GetPrinter(hPrinter, 2, PrinterInfo(1), ByteBuf, _
     BytesNeeded)
   
   'Check for errors.
   If Result = 0 Then
      'Determine the error that occurred.
      LastError = err.LastDllError()
      
      'Display error message, close printer, and exit sub.
      CheckPrinter = "Couldn't get Printer Status!  Error = " _
         & LastError
      ClosePrinter hPrinter
      Exit Function
   End If

   'Copy contents of printer status byte array into a
   'PRINTER_INFO_2 structure to separate the individual elements.
   CopyMemory PI2, PrinterInfo(1), Len(PI2)
   
   'Check if printer is in ready state.
   PrinterStr = CheckPrinterStatus(PI2.Status)
   
'''''   'Add printer name, driver, and port to list.
'''''   PrinterStr = PrinterStr & "Printer Name = " & _
'''''     GetString(PI2.pPrinterName) & vbCrLf
'''''   PrinterStr = PrinterStr & "Printer Driver Name = " & _
'''''     GetString(PI2.pDriverName) & vbCrLf
'''''   PrinterStr = PrinterStr & "Printer Port Name = " & _
'''''     GetString(PI2.pPortName) & vbCrLf
   
   'Call API to get size of buffer that is needed.
   Result = EnumJobs(hPrinter, 0&, &HFFFFFFFF, 2, ByVal 0&, 0&, _
      BytesNeeded, NumJI2)
   
   'Check if there are no current jobs, and then display appropriate message.
   If BytesNeeded = 0 Then
      JobStr = "No Print Jobs!"
   Else
      'Redim byte array to hold info about print job.
      ReDim JobInfo(0 To BytesNeeded)
      
      'Call API to get print job info.
      Result = EnumJobs(hPrinter, 0&, &HFFFFFFFF, 2, JobInfo(0), _
        BytesNeeded, ByteBuf, NumJI2)
      
      'Check for errors.
      If Result = 0 Then
         'Get and display error, close printer, and exit sub.
         LastError = err.LastDllError
         CheckPrinter = " > EnumJobs Failed on second call! <  Error = " _
            & LastError
         ClosePrinter hPrinter
         Exit Function
      End If
      
      
      'Copy contents of print job info byte array into a
      'JOB_INFO_2 structure to separate the individual elements.
      For i = 0 To NumJI2 - 1   ' Loop through jobs and walk the buffer
          CopyMemory JI2, JobInfo(i * Len(JI2)), Len(JI2)
             
''''          ' List info available on Jobs.
''''          Text2 = Text2 & "Job ID" & vbTab & JI2.JobId & vbCrLf
''''          Text2 = Text2 & "Name Of Printer" & vbTab & _
''''            GetString(JI2.pPrinterName) & vbCrLf
''''          Text2 = Text2 & "Name Of Machine That Created Job" & vbTab & _
''''            GetString(JI2.pMachineName) & vbCrLf
''''          Text2 = Text2 & "Print Job Owner's Name" & vbTab & _
''''            GetString(JI2.pUserName) & vbCrLf
''''          Text2 = Text2 & "Name Of Document" & vbTab & GetString(JI2.pDocument)
''''          Text2 = Text2 & "Name Of User To Notify" & vbTab & _
''''            GetString(JI2.pNotifyName) & vbCrLf
''''          Text2 = Text2 & "Type Of Data" & vbTab & GetString(JI2.pDatatype)
''''          Text2 = Text2 & "Print Processor" & vbTab & _
''''            GetString(JI2.pPrintProcessor) & vbCrLf
''''          Text2 = Text2 & "Print Processor Parameters" & vbTab & _
''''            GetString(JI2.pParameters) & vbCrLf
''''          Text2 = Text2 & "Print Driver Name" & vbTab & _
''''            GetString(JI2.pDriverName) & vbCrLf
''''          Text2 = Text2 & "Print Job 'P' Status" & vbTab & _
''''            GetString(JI2.pStatus) & vbCrLf
''''          Text2 = Text2 & "Print Job Status" & vbTab & JI2.Status & vbCrLf
''''          Text2 = Text2 & "Print Job Priority" & vbTab & JI2.Priority & vbCrLf
''''          Text2 = Text2 & "Position in Queue" & vbTab & JI2.Position & vbCrLf
''''          Text2 = Text2 & "Earliest Time Job Can Be Printed" & vbTab & _
''''            JI2.StartTime & vbCrLf
''''          Text2 = Text2 & "Latest Time Job Will Be Printed" & vbTab & _
''''            JI2.UntilTime & vbCrLf
''''          Text2 = Text2 & "Total Pages For Entire Job" & vbTab & JI2.TotalPages & vbCrLf
''''          Text2 = Text2 & "Size of Job In Bytes" & vbTab & JI2.Size & vbCrLf
''''          'Because of a bug in Windows NT 3.51, the time member is not set correctly.
''''          'Therefore, do not use the time member on Windows NT 3.51.
''''          Text2 = Text2 & "Elapsed Print Time" & vbTab & JI2.time & vbCrLf
''''          Text2 = Text2 & "Pages Printed So Far" & vbTab & JI2.PagesPrinted & vbCrLf
             
          'Display basic job status info.
          JobStr = JobStr & "Job ID = " & JI2.JobId & _
             vbCrLf & "Total Pages = " & JI2.TotalPages & vbCrLf
            
            JobQuantity = JobQuantity + 1
            
            TempStr = ""   'Clear
          'Check for a ready state.
          If JI2.pStatus = 0& Then   ' If pStatus is Null, check Status.
            If JI2.Status = 0 Then
               TempStr = TempStr & "Ready!  " & vbCrLf
            Else  'Check for the various print job states.
               If (JI2.Status And JOB_STATUS_SPOOLING) Then
                  TempStr = TempStr & "Spooling  "
               End If
               
               If (JI2.Status And JOB_STATUS_OFFLINE) Then
                  TempStr = TempStr & "Off line  "
               End If
               
               If (JI2.Status And JOB_STATUS_PAUSED) Then
                  TempStr = TempStr & "Paused  "
               End If
               
               If (JI2.Status And JOB_STATUS_ERROR) Then
                  TempStr = TempStr & "Error  "
               End If
               
               If (JI2.Status And JOB_STATUS_PAPEROUT) Then
                  TempStr = TempStr & "Paper Out  "
               End If
               
               If (JI2.Status And JOB_STATUS_PRINTING) Then
                  TempStr = TempStr & "Printing  "
               End If
               
               If (JI2.Status And JOB_STATUS_USER_INTERVENTION) Then
                  TempStr = TempStr & "User Intervention Needed  "
               End If
               
               If Len(TempStr) = 0 Then
                  TempStr = "Unknown Status of " & JI2.Status
               End If
            End If
        Else
            ' Dereference pStatus.
            TempStr = PtrCtoVbString(JI2.pStatus)
        End If
          
          'Report the Job status.
          JobStr = JobStr & TempStr & vbCrLf
          'Debug.Print JobStr & TempStr
      Next i
   End If
   
   'Close the printer handle.
   ClosePrinter hPrinter
End Function



