VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl CallerIDMonitor 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   Picture         =   "CallerIDMonitor.ctx":0000
   ScaleHeight     =   990
   ScaleWidth      =   1200
   Begin VB.TextBox TLen 
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Text            =   "0"
      Top             =   90
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   645
      Top             =   60
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   57600
   End
End
Attribute VB_Name = "CallerIDMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim CallerIdPort As String
Dim LineNumber As Byte
Dim LineNum As Integer
Dim Flash As Byte
Dim divider_string As String
Dim Timer_Counter As Byte
Dim main_string, recieve_string
Dim s, Number, Inputstr As String
'Event Declarations:
Event CallerIDDetect(Line As Byte, Number As String)

Private Sub divide_string(ByRef recieve_string)                  ''—Ê«· Å—œ«“‘ «ÿ·«⁄«  œ—Ì«ﬁ Ì «“ œ” ê«Â ﬂ«·—¬ÌœÌ
Dim end_string_flag, recieve_text, at_position
Dim current_text As String
Dim number_string As String
end_string_flag = True
number_string = 0
While end_string_flag                                            ''«œ«„Â ﬂ«— Å—œ«“‘ «ÿ·«⁄«   « “„«‰Ì ﬂÂ ﬂ«—«ﬂ — Ãœ« ﬂ‰‰œÂ @ œ— —‘ Â ÊÃÊœ œ«‘ Â »«‘œ
        at_position = InStr(recieve_string, divider_string)
        current_text = Mid$(recieve_string, 1, at_position)
         If current_text <> "" Then
            current_text = Get_CallerID(current_text)           '' «—”«·  ﬁ”„ Ì «“ —‘ Â ﬂÂ »Â Ìﬂ Ãœ« ﬂ‰‰œÂ —”ÌœÂ «”  »Â  «»⁄ Å—œ«“‘
            If current_text <> "" Then
                If Val(TLen.Text) > 0 Then current_text = Right(current_text, Len(current_text) - (Val(TLen.Text)))
                current_text = ProccessNumber(current_text)       '«—”«· ‘„«—Â  „«” »œ”  ¬„œÂ ÃÂ  Å—œ«“‘ »—«Ì Õ–› ÅÌ‘ ‘„«—Â Â«Ì «Õ „«·Ì - ﬂœ ‘Â—
                RaiseEvent CallerIDDetect(LineNumber, current_text)

            End If
         End If
         recieve_string = Mid$(recieve_string, at_position + 1)
      If (InStr(recieve_string, divider_string) <> 0) Then       ''«œ«„Â ﬂ«— Å—œ«“‘ «ÿ·«⁄«   « “„«‰Ì ﬂÂ ﬂ«—«ﬂ — Ãœ« ﬂ‰‰œÂ @ œ— —‘ Â ÊÃÊœ œ«‘ Â »«‘œ
                end_string_flag = True
            Else
                end_string_flag = False
       End If
Wend
End Sub

Private Sub Timer1_Timer()
    recieve_string = MSComm1.Input  ' ' ŒÊ«‰œ‰ »«›— ﬂ‰ —· «„ «” ﬂ«„
    If (recieve_string <> "") Then
      '  Inputstr = recieve_string  ############
        Timer_Counter = 0
        main_string = main_string & recieve_string   ' '  ﬁ—«— œ«œ‰ „ﬁœ«— ÃœÌœ œ— «‰ Â«Ì „ﬁ«œ— «Õ „«·Ì »«ﬁÌ „«‰œÂ «“ ﬁ»·
        recieve_string = ""
        Flash = 0
    Else
        If main_string <> "" Then '(read_flag) Then
                Timer_Counter = Timer_Counter + 1
                Call divide_string(main_string)
                If Timer_Counter > 10 Then
                    main_string = ""
                    Timer_Counter = 0
                End If
         End If
    End If
End Sub
Private Sub RemoveZero(ByRef Str As String)
If Len(Str) = 0 Then Exit Sub
Do While (Asc(Left(Str, 1)) < 48)
    Str = Right(Str, Len(Str) - 1)
    If (Len(Str) = 0) Then Exit Do
Loop
While Left(Str, 1) = "0"
    Str = Right(Str, Len(Str) - 1)
Wend

End Sub
Private Function Get_CallerID(ByVal strin As String) As String    ' «»⁄ Å—œ«“‘ “Ì— —‘ Â œ—Ì«› Ì ÃÂ  Å«·«Ì‘ ‘„«—Â  „«” «“ ¬‰
Dim CallerID As String
CallerID = strin
If InStr(1, CallerID, "T") > 0 Then Exit Function
If InStr(1, CallerID, "P") > 0 Then Exit Function

Dim num1 As String
Dim num2 As String
Dim Start_Line As Byte
Dim Line As Byte
If Len(CallerID) < 7 Then Exit Function
Call RemoveZero(CallerID)
Line = Val(Mid(CallerID, 2, 1))
LineNumber = Line
If divider_string = "@" Then
    CallerID = Right(CallerID, Len(CallerID) - 3)
Else
    CallerID = Right(CallerID, Len(CallerID) - 2)
End If
'##Arya
If InStr(LCase(CallerID), "callerid:") > 0 Then
    CallerID = LTrim(Mid(CallerID, InStr(1, LCase(CallerID), "callerid:") + 9))
End If
'##
Call RemoveZero(CallerID)
'--- Detect Country Code ----
Do While (Asc(Left(CallerID, 1)) > 47) And (Asc(Left(CallerID, 1)) < 58)
    num1 = num1 & Left(CallerID, 1)
    CallerID = Right(CallerID, Len(CallerID) - 1)
    If Len(CallerID) = 0 Then Exit Do
Loop
Call RemoveZero(CallerID)
If Len(CallerID) < 8 Then
    Get_CallerID = num1
    Exit Function
End If

If Left(CallerID, 1) = "L" Then
    LineNumber = Line
    CallerID = Right(CallerID, Len(CallerID) - 3)
End If
While (Asc(Left(CallerID, 1)) > 47) And (Asc(Left(CallerID, 1)) < 58)
    num2 = num2 & Left(CallerID, 1)
    CallerID = Right(CallerID, Len(CallerID) - 1)
Wend
Call RemoveZero(CallerID)
   Get_CallerID = num2
   Exit Function
End Function

Private Function ProccessNumber(ByVal strin As String) As String  ' «»⁄ Õ–› ÅÌ‘ ‘„«—Â Â«Ì «Õ „«·Ì «“ —‘ Â ‰Â«ÌÌ'
If Left(strin, 2) = "98" Then strin = Right(strin, Len(strin) - 2)
Call RemoveZero(strin)
'--- Detect Mobile Call ----
If Left(strin, 1) = "9" Then
    ProccessNumber = "0" & strin
Else
'    If Left(strin, 3) = "511" Then
'        strin = Right(strin, Len(strin) - 3)
'    ElseIf Left(strin, 3) = "21" Then
'        strin = Right(strin, Len(strin) - 2)
'    End If
    ProccessNumber = strin
End If
End Function
Public Property Get PortNumber() As Integer
    PortNumber = MSComm1.CommPort
End Property
Public Property Let PortNumber(ByVal New_PortNumber As Integer)
    MSComm1.CommPort() = New_PortNumber
    PropertyChanged "PortNumber"
End Property
Public Property Get Baudrate() As Long

End Property
Public Property Let Baudrate(ByVal New_Baudrate As Long)
  '  MSComm1.PortOpen = PropBag.ReadProperty("OpenPort", New_Baudrate, n, 8, 1)
    MSComm1.Settings = New_Baudrate & ", n, 8, 1"
    PropertyChanged "Baudrate"
End Property
Private Sub UserControl_Initialize()
divider_string = "@" ' ﬂ«—«ﬂ —  ⁄ÌÌ‰ ﬂ‰‰œÂ Å«Ì«‰ «ÿ·«⁄«  «—”«·Ì «“ ﬂ«·—¬ÌœÌ
End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MSComm1.CommPort = PropBag.ReadProperty("PortNumber", 1)
    TLen.Text = PropBag.ReadProperty("RemoveLen", "4")
    'MSComm1.PortOpen = PropBag.ReadProperty("OpenPort", 57600, n, 8, 1)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("PortNumber", MSComm1.CommPort, 1)
    Call PropBag.WriteProperty("RemoveLen", TLen.Text, 4)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TLen,TLen,-1,Text
Public Property Get RemoveLen() As String
    RemoveLen = TLen.Text
End Property

Public Property Let RemoveLen(ByVal New_RemoveLen As String)
    TLen.Text() = New_RemoveLen
    PropertyChanged "RemoveLen"
End Property
'
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0
'
'Public Function OpenPort(Status As Boolean) As Boolean
'
'        MSComm1.PortOpen = Status
'        If Status = True Then
'            Timer1.Enabled = True
'        Else
'            Timer1.Enabled = False
'        End If
'End Function
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSComm1,MSComm1,-1,Output
Public Property Get OpenPort() As Boolean
    OpenPort = MSComm1.PortOpen
End Property

Public Property Let OpenPort(ByVal New_OpenPort As Boolean)
    On Error GoTo ErrHandler
    MSComm1.PortOpen = New_OpenPort
    MSComm1.Output = Chr$(252) + Chr$(112) + Chr$(114) + Chr$(116) + Chr$(6) + Chr$(108) + Chr$(253)    ' INIT Protocol #6    If New_OpenPort = True Then
    If New_OpenPort = True Then
       Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    
    End If
    PropertyChanged "OpenPort"
Exit Property
ErrHandler:
    MsgBox "Œÿ« œ— »«“ò—œ‰ ÅÊ—  ò«·—¬Ì œÌ  - " & err.Description
End Property


