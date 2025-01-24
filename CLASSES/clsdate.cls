VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Type Date_Structure
    year As Integer
    Month As Integer
    day As Integer
    day_week As Integer
End Type
Dim FarDate1

Public Function Find_DayOfWeek(Week_Day As String) As String
If clsArya.MiladiDate = 0 Then
    Select Case Week_Day
        Case "1"
            Find_DayOfWeek = "ÔäÈå"
        Case "2"
            Find_DayOfWeek = "íß ÔäÈå"
        Case "3"
            Find_DayOfWeek = "Ïæ ÔäÈå"
        Case "4"
            Find_DayOfWeek = "Óå ÔäÈå"
        Case "5"
            Find_DayOfWeek = "åÇÑ ÔäÈå"
        Case "6"
            Find_DayOfWeek = "äÌ ÔäÈå"
        Case "7"
            Find_DayOfWeek = "ÌãÚå"
    End Select
Else
    Select Case Week_Day
        Case "1"
            Find_DayOfWeek = "Saturday"
        Case "2"
            Find_DayOfWeek = "Sunday"
        Case "3"
            Find_DayOfWeek = "Monday"
        Case "4"
            Find_DayOfWeek = "Tuesday"
        Case "5"
            Find_DayOfWeek = "Wednesday"
        Case "6"
            Find_DayOfWeek = "Thursday"
        Case "7"
            Find_DayOfWeek = "Friday"
    End Select
End If
End Function
Public Function Find_DayOfWeekShamsi(strDate As String) As String
    Find_DayOfWeekShamsi = Miladi(strDate)
    Find_DayOfWeekShamsi = Find_DayOfWeek(Weekday(Find_DayOfWeekShamsi, vbSaturday))
End Function

Public Function Km(ByVal ym As Integer) As Integer
    Km = (Int((ym / 100)) - Int((ym / 400)) - 2) * IIf(ym > 1581, -1, 0) + Int(ym / 4)
End Function

Public Function Ki(ByVal yi As Integer) As Integer
    Dim Y As Integer
    Dim Y1  As Integer
    Dim Y2  As Integer
    Dim y3  As Integer
    Dim y4  As Integer
    Y = yi - 508
    Y1 = Int(Y / 128)
    Y2 = Y Mod 128
    y3 = Int(Y2 / 33)
    y4 = Y2 Mod 33
    Ki = (Y1 * 31 + y3 * 8 + y4 / 4 - Y2 / 127 - y4 / 32)
'    Debug.Print "yi = " & yi
'    Debug.Print "y = " & y
'    Debug.Print "y1 = " & y1
'    Debug.Print "y2 = " & y2
'    Debug.Print "y3 = " & y3
'    Debug.Print "y4 = " & y4
'    Debug.Print "Ki = " & Ki
'    Debug.Print (y1 * 31 + y3 * 8 + y4 / 4 - y2 / 127 - y4 / 32)
End Function

Public Function Rsm(ByVal yi As Integer) As Integer
   Rsm = Ki(yi - 1) - Km(yi + 620) + 355
End Function


Public Function table2(ByVal iMonth As Integer) As Integer
    table2 = 30 * iMonth + IIf(iMonth > 6, 6, iMonth)
End Function

Public Function table1(ByVal iYear As Integer, ByVal iMonth As Integer) As Integer
    Dim D(11) As Integer
    D(0) = 0
    D(1) = 31
    D(2) = 59
    D(3) = 90
    D(4) = 120
    D(5) = 151
    D(6) = 181
    D(7) = 212
    D(8) = 243
    D(9) = 273
    D(10) = 304
    D(11) = 334
    If iYear = 1 Then
        table1 = D(iMonth) + IIf(iMonth >= 2, 1, 0)
    Else
        table1 = D(iMonth)
    End If
End Function

Public Function Miladi(ByVal strDate As String)
'    Dim booLongdate As Boolean
'    Dim i As Integer
'    Dim r As Integer, _
'        DayOfYear As Integer, _
'        r2 As Integer, _
'        mk As Integer
'    Dim ir As Date_Structure, _
'        mi As Date_Structure
'
'        ir.year = Val(Left(strDate, 4))
'        ir.Month = Val(Mid(strDate, 6, 2))
'    ir.day = Val(Right(strDate, 2))
'    mi.year = ir.year + 621
'    r = Rsm(ir.year)
'    DayOfYear = ir.day + table2(ir.Month - 1)
'    r2 = r + DayOfYear + 1
'    mk = Km(mi.year) - Km(mi.year - 1)
'    If (r2 > 365 + mk) Then
'        mi.year = mi.year + 1
'          r2 = r2 - 365 - mk
'          mk = Km(mi.year) - Km(mi.year - 1)
'    End If
'    For i = 1 To 12
'        If r2 <= table1(mk, i - 1) Then Exit For
'    Next
'    i = i - 1
'    mi.Month = i
'    mi.day = r2 - table1(mk, i - 1)
'    Miladi = Format(mi.year, "0###") & "/" & Format(mi.Month, "0#") & "/" & Format(mi.day, "0#")

    Miladi = FarDate1.ShamsiToMiladi(Left(strDate, 4) & Mid(strDate, 6, 2) & Right(strDate, 2))
End Function
Private Sub Class_Initialize()
   ' Set FarDate1 = CreateObject("ADPDatePicker.FarDate")
    Set FarDate1 = CreateObject("FarDate1.FarDate")
End Sub

Private Sub Class_Terminate()
     Set FarDate1 = Nothing
End Sub
Public Function shamsi(sDate As Date) As String
    If clsArya.MiladiDate = 0 Then
        shamsi = FarDate1.MiladiToShamsi(sDate)
        shamsi = Left(shamsi, 4) & "/" & Mid(shamsi, 5, 2) & "/" & Right(shamsi, 2)
'        shamsi = mdifrm.FarDate1.MiladiToShamsi(sDate)
'        shamsi = Left(shamsi, 4) & "/" & Mid(shamsi, 5, 2) & "/" & Right(shamsi, 2)
    Else
        shamsi = CStr(year(sDate)) + "/" + Format(CStr(Month(sDate)), "00") + "/" + Format(CStr(day(sDate)), "00")
    End If
End Function
Public Function shamsiAddedDate(sDate As Date, AddedDate As Double) As String
    shamsiAddedDate = FarDate1.MiladiToShamsi(DateAdd("d", AddedDate, sDate))
    shamsiAddedDate = Left(shamsiAddedDate, 4) & "/" & Mid(shamsiAddedDate, 5, 2) & "/" & Right(shamsiAddedDate, 2)
    'shamsiAddedDate = Format(ir.year, "0###") & "/" & Format(ir.Month, "0#") & "/" & Format(ir.day, "0#")
End Function

Public Function validdate(ByVal strDate As String) As Boolean
    Dim intYear As Integer, intMonth As Integer, intDay As Integer
    If Len(strDate) = 8 Then
        strDate = "13" & strDate
    End If
    If Len(strDate) <> 10 Or Mid(strDate, 5, 1) <> "/" Or Mid(strDate, 8, 1) <> "/" Then
        Exit Function
    End If
    intDay = Val(Right(strDate, 2))
    intMonth = Val(Mid(strDate, 6, 2))
    intYear = Val(Left(strDate, 4))
    If (intDay >= 1 And intDay <= 31 And intMonth >= 1 And intMonth <= 12) And (intMonth <= 6 Or intDay <= 30) Then
        validdate = True
    End If
End Function

Public Function DateRemain(ByVal startDate As String, ByVal EndDate As String) As Long
    Dim intStartYear As Long, intStartMonth As Long, intStartDay As Long
    Dim intEndYear As Long, intEndMonth As Long, intEndDay As Long
    Dim startDateDays As Long, EndDateDays As Long
    
    intStartDay = Val(Right(startDate, 2))
    intStartMonth = Val(Mid(startDate, 3, 2))
    intStartYear = Val(Left(startDate, 2))
    
    startDateDays = 30 * intStartMonth + IIf(intStartMonth > 6, 6, intStartMonth)
    startDateDays = 365 * intStartYear + startDateDays + intStartDay
    
    intEndDay = Val(Right(EndDate, 2))
    intEndMonth = Val(Mid(EndDate, 3, 2))
    intEndYear = Val(Left(EndDate, 2))
    EndDateDays = 30 * intEndMonth + IIf(intEndMonth > 6, 6, intEndMonth)
    EndDateDays = 365 * intEndYear + EndDateDays + intEndDay

    DateRemain = EndDateDays - startDateDays
    
End Function

