Attribute VB_Name = "ModDataBaseHandler"

Type Parameter
 
    Name As String
    DataType As DataTypeEnum
    Direction As ParameterDirectionEnum
    Size As Long
    Value As Variant

End Type

Public Function RunStoredProcedure2RecordSet(SpName As String, Optional AdoConnection As ADODB.Connection) As ADODB.Recordset
    
    Dim MyAdoCommand As New ADODB.Command
    Dim MyAdoRecordset As New ADODB.Recordset
    
    If AdoConnection Is Nothing Then Set AdoConnection = PosConnection
    
    If AdoConnection.ConnectionString = "" Then
        AdoConnection.ConnectionString = strConnectionString
    End If
    If AdoConnection.State <> 1 Then
        AdoConnection.Open
    End If
    
    With MyAdoCommand
        
        .ActiveConnection = AdoConnection
        .CommandType = adCmdStoredProc
        .CommandText = SpName
        .CommandTimeout = 180
    End With
    
    Set RunStoredProcedure2RecordSet = New ADODB.Recordset
    RunStoredProcedure2RecordSet.CursorType = adOpenDynamic
    
'    MyAdoRecordset.CursorType = adOpenDynamic
'    Set MyAdoRecordset = MyAdoCommand.Execute
    Set RunStoredProcedure2RecordSet = MyAdoCommand.Execute

    Set MyAdoCommand = Nothing
'    Set RunStoredProcedure2RecordSet = MyAdoRecordset
    
'     If AdoConnection.State = 1 Then
'        AdoConnection.Close
'    End If
End Function

Function RunParametricStoredProcedure2Rec(SpName As String, Parameter() As Parameter, Optional AdoConnection As ADODB.Connection) As ADODB.Recordset

    Dim MyAdoCommand As New ADODB.Command
    Dim index As Integer
    
    If AdoConnection Is Nothing Then Set AdoConnection = PosConnection
    
    If AdoConnection.ConnectionString = "" Then
        AdoConnection.ConnectionString = strConnectionString
    End If

    If AdoConnection.State <> 1 Then
        AdoConnection.Open
    End If

    With MyAdoCommand
        .ActiveConnection = AdoConnection
        .CommandType = adCmdStoredProc
        .CommandText = SpName
        .CommandTimeout = 180
    End With
    
    For index = LBound(Parameter) To UBound(Parameter)
            MyAdoCommand.Parameters.Append MyAdoCommand.CreateParameter(Parameter(index).Name, Parameter(index).DataType, Parameter(index).Direction, Parameter(index).Size, Parameter(index).Value)
    Next index
    
    Set RunParametricStoredProcedure2Rec = New ADODB.Recordset
    
  Set RunParametricStoredProcedure2Rec = MyAdoCommand.Execute
    
    Set MyAdoCommand = Nothing
End Function


 Function RunNonParametricStoredProcedure(SpName As String, Optional AdoConnection As ADODB.Connection)
    
    Dim MyAdoCommand As New ADODB.Command
    
    If AdoConnection Is Nothing Then Set AdoConnection = PosConnection
    
    If AdoConnection.ConnectionString = "" Then
        AdoConnection.ConnectionString = strConnectionString
    End If
    If AdoConnection.State <> 1 Then
        AdoConnection.Open
    End If
    
    With MyAdoCommand
        .ActiveConnection = AdoConnection
        .CommandType = adCmdStoredProc
        .CommandText = SpName
        .CommandTimeout = 180
        .Execute
    End With

    Set MyAdoCommand = Nothing

End Function

 Function RunParametricStoredProcedure(SpName As String, Parameter() As Parameter, Optional AdoConnection As ADODB.Connection) As Long

    Dim MyAdoCommand As New ADODB.Command
    Dim index As Integer
    
    If AdoConnection Is Nothing Then Set AdoConnection = PosConnection
    
    If AdoConnection.ConnectionString = "" Then
        AdoConnection.ConnectionString = strConnectionString
    End If
    If AdoConnection.State <> 1 Then
        AdoConnection.Open
    End If

    With MyAdoCommand
        .ActiveConnection = AdoConnection
        .CommandType = adCmdStoredProc
        .CommandText = SpName
        .CommandTimeout = 180
    End With
    
    For index = LBound(Parameter) To UBound(Parameter)
            MyAdoCommand.Parameters.Append MyAdoCommand.CreateParameter(Parameter(index).Name, Parameter(index).DataType, Parameter(index).Direction, Parameter(index).Size, Parameter(index).Value)
    Next index
    
    MyAdoCommand.Execute
    If MyAdoCommand.Parameters(MyAdoCommand.Parameters.Count - 1).Direction = adParamInputOutput Or MyAdoCommand.Parameters(MyAdoCommand.Parameters.Count - 1).Direction = adParamOutput Then
        RunParametricStoredProcedure = Val(MyAdoCommand.Parameters(MyAdoCommand.Parameters.Count - 1).Value)
    End If
    Set MyAdoCommand = Nothing

End Function

 Function RunParametricStoredProcedureReturnValue(SpName As String, Parameter() As Parameter, Optional AdoConnection As ADODB.Connection) As Long

    Dim MyAdoCommand As New ADODB.Command
    Dim index As Integer

    If AdoConnection Is Nothing Then Set AdoConnection = PosConnection
    
    If AdoConnection.ConnectionString = "" Then
        AdoConnection.ConnectionString = strConnectionString
    End If
    If AdoConnection.State <> 1 Then
        AdoConnection.Open
    End If

    With MyAdoCommand
        .ActiveConnection = AdoConnection
        .CommandType = adCmdStoredProc
        .CommandText = SpName
        .CommandTimeout = 180
    End With
    
    For index = LBound(Parameter) To UBound(Parameter)
            MyAdoCommand.Parameters.Append MyAdoCommand.CreateParameter(Parameter(index).Name, Parameter(index).DataType, Parameter(index).Direction, Parameter(index).Size, Parameter(index).Value)
    Next index
    
    MyAdoCommand.Parameters.Append MyAdoCommand.CreateParameter("ReturnValue", adInteger, adParamReturnValue, 4)
    
    MyAdoCommand.Execute
    
    RunParametricStoredProcedureReturnValue = Val(MyAdoCommand.Parameters.Item("ReturnValue"))

    Set MyAdoCommand = Nothing

End Function


 Function RunParametricStoredProcedure2String(SpName As String, Parameter() As Parameter, Optional AdoConnection As ADODB.Connection) As String

    Dim MyAdoCommand As New ADODB.Command
    Dim index As Integer

    If AdoConnection Is Nothing Then Set AdoConnection = PosConnection

    If AdoConnection.ConnectionString = "" Then
        AdoConnection.ConnectionString = strConnectionString
    End If
    If AdoConnection.State = 0 Then
        AdoConnection.Open
    End If

    With MyAdoCommand
        .ActiveConnection = AdoConnection
        .CommandType = adCmdStoredProc
        .CommandText = SpName
        .CommandTimeout = 180
    End With
    
    For index = LBound(Parameter) To UBound(Parameter)
            MyAdoCommand.Parameters.Append MyAdoCommand.CreateParameter(Parameter(index).Name, Parameter(index).DataType, Parameter(index).Direction, Parameter(index).Size, Parameter(index).Value)
    Next index
    
    MyAdoCommand.Execute
    If MyAdoCommand.Parameters(MyAdoCommand.Parameters.Count - 1).Direction = adParamInputOutput Or MyAdoCommand.Parameters(MyAdoCommand.Parameters.Count - 1).Direction = adParamOutput Then
        RunParametricStoredProcedure2String = CStr(MyAdoCommand.Parameters(MyAdoCommand.Parameters.Count - 1).Value)
    End If
    Set MyAdoCommand = Nothing

End Function

 Function GenerateInputParameter(ParameterName As String, ParameterDatatype As DataTypeEnum, ParameterSize As Long, ParameterValue As Variant) As Parameter
       
    GenerateInputParameter.DataType = ParameterDatatype
    GenerateInputParameter.Direction = adParamInput
    GenerateInputParameter.Name = ParameterName
    GenerateInputParameter.Size = ParameterSize
    GenerateInputParameter.Value = ParameterValue
        
End Function

Function GenerateOutputParameter(ParameterName As String, ParameterDatatype As DataTypeEnum, ParameterSize As Long) As Parameter
      
    GenerateOutputParameter.DataType = ParameterDatatype
    GenerateOutputParameter.Direction = adParamOutput
    GenerateOutputParameter.Name = ParameterName
    GenerateOutputParameter.Size = ParameterSize
    GenerateOutputParameter.Value = 0

End Function

 Function GenerateInputOutputParameter(ParameterName As String, ParameterDatatype As DataTypeEnum, ParameterSize As Long, ParameterValue As Variant) As Parameter
       
    GenerateInputOutputParameter.DataType = ParameterDatatype
    GenerateInputOutputParameter.Direction = adParamInputOutput
    GenerateInputOutputParameter.Name = ParameterName
    GenerateInputOutputParameter.Size = ParameterSize
    GenerateInputOutputParameter.Value = ParameterValue
        
End Function

Function GenerateDetailsString(ByRef DetailsString As String, amount As String, GoodCode As String, FeeUnit As String, Discount As String, Rate As String, ChairName As String, ExpireDate As String, InventoryNo As String, DestInventoryNo As String, ServePlace As String, Optional Differences As String = "", Optional Seller As String = "") As String
    
    
    GenerateDetailsString = DetailsString & amount & ";" & GoodCode & ";" & FeeUnit & ";" & Discount & ";" & Rate & ";" & ChairName & ";" & ExpireDate & ";" & InventoryNo & ";" & DestInventoryNo & ";" & Seller & ";" & ServePlace   '& "/"
    
    If Differences <> "" Then
    
        Dim ArrDifferences() As String
        Dim i As Integer
        
        ArrDifferences = Split(Differences, ";")
        For i = LBound(ArrDifferences) To UBound(ArrDifferences)
            GenerateDetailsString = GenerateDetailsString & ";" & ArrDifferences(i)
        
        Next i
    End If
    GenerateDetailsString = GenerateDetailsString & "/"
End Function
Function GenerateDetailsString3(ByRef DetailsString As String, amount As String, GoodCode As String, FeeUnit As String, Discount As String, Rate As String, ChairName As String, ExpireDate As String, InventoryNo As String, DestInventoryNo As String, ServePlace As String, Optional Differences As String) As String
    
    
    GenerateDetailsString3 = DetailsString & amount & ";" & GoodCode & ";" & FeeUnit & ";" & Discount & ";" & Rate & ";" & ChairName & ";" & ExpireDate & ";" & InventoryNo & ";" & DestInventoryNo & ";" & ServePlace    '& "/"
    If Differences = "" Then
        GenerateDetailsString3 = GenerateDetailsString3 & "/"
    Else
    
        Dim ArrDifferences() As String
        Dim i As Integer
        
        ArrDifferences = Split(Differences, ";")
        For i = LBound(ArrDifferences) To UBound(ArrDifferences)
            GenerateDetailsString3 = GenerateDetailsString3 & ";" & ArrDifferences(i)
        
        Next i
        GenerateDetailsString3 = GenerateDetailsString3 & "/"
    End If
    
End Function

Public Function GenerateDetailsString2(ByRef DetailsString As String, AccountYear As String, Branch As String, DocumentId As String, RowId As String, kolId As String, MoeinId As String, TafsiliId As String, RowDes As String, Bedehkar As Long, Bestankar As Long, kind As Integer, SaveDate As String, UserID As Integer) As String
    GenerateDetailsString2 = DetailsString & "/$/" & AccountYear & "/^/" & Branch & "/^/" & DocumentId & "/^/" & RowId & "/^/" & kolId & "/^/" & MoeinId & "/^/" & TafsiliId & "/^/" & RowDes & "/^/" & Bedehkar & "/^/" & Bestankar & "/^/" & kind & "/^/" & SaveDate & "/^/" & UserID & "/^/"
End Function

Public Function GenerateDetailsStringFactorReceived(ByRef DetailsString As String, c1 As String, c2 As String, c3 As String, c4 As String, c5 As String, c6 As String, c7 As String, c8 As String, c9 As String, c10 As String) As String
    GenerateDetailsStringFactorReceived = DetailsString & "/$/" & c1 & "/^/" & c2 & "/^/" & c3 & "/^/" & c4 & "/^/" & c5 & "/^/" & c6 & "/^/" & c7 & "/^/" & c8 & "/^/" & c9 & "/^/" & c10 & "/^/"
End Function

Function RunMPSP2Rec(SpName As String, Parameter() As Parameter, Optional AdoConnection As ADODB.Connection) As ADODB.Recordset

    Dim index As Integer

    If AdoConnection Is Nothing Then Set AdoConnection = PosConnection

    If AdoConnection.ConnectionString = "" Then
        AdoConnection.ConnectionString = strConnectionString
    End If
    If AdoConnection.State <> 1 Then
        AdoConnection.Open
    End If
    
    Dim strSource As String
    
    For index = LBound(Parameter) To UBound(Parameter)
            strSource = strSource & Parameter(index).Value & " , "
    Next index
    
    If strSource <> "" Then
        strSource = Left(strSource, Len(strSource) - 2)
    End If
    strSource = " Exec " & SpName & " " & strSource
    Set RunMPSP2Rec = New ADODB.Recordset
    
    RunMPSP2Rec.Open strSource, AdoConnection, adOpenDynamic, adLockOptimistic, adCmdText
    
    Set MyAdoCommand = Nothing

End Function

Public Function GenerateDetailsStringAccount(ByRef DetailsString As String, AccountYear As String, Branch As String, DocumentId As String, RowId As String, kolId As String, MoeinId As String, TafsiliId As String, RowDes As String, Bedehkar As Long, Bestankar As Long, kind As Integer, SaveDate As String, UserID As Integer, CheckNo As String, CheckDate As String) As String
    GenerateDetailsStringAccount = DetailsString & "/$/" & AccountYear & "/^/" & Branch & "/^/" & DocumentId & "/^/" & RowId & "/^/" & kolId & "/^/" & MoeinId & "/^/" & TafsiliId & "/^/" & RowDes & "/^/" & Bedehkar & "/^/" & Bestankar & "/^/" & kind & "/^/" & SaveDate & "/^/" & UserID & "/^/" & CheckNo & "/^/" & CheckDate & "/^/"
End Function
Public Function HamyarGenerateDetailsStringAccount(ByRef DetailsString As String, No As String, Satr As String, Kol As String, M1 As String, M2 As String, Descs As String, Bedeh As Long, Bestan As Long, Dates As String) As String
    HamyarGenerateDetailsStringAccount = DetailsString & "/$/" & No & "/^/" & Satr & "/^/" & Kol & "/^/" & M1 & "/^/" & M2 & "/^/" & Descs & "/^/" & Bedeh & "/^/" & Bestan & "/^/" & Dates & "/^/"
End Function

Public Function GenerateDetailsStringReportGenarator(ByRef DetailsString As String, c0 As String, c1 As String, c2 As String, c3 As String, c4 As String, c5 As String, c6 As String, c7 As String, c8 As String, c9 As String, c10 As String, c11 As String, c12 As String, c13 As String) As String
    GenerateDetailsStringReportGenarator = DetailsString & "/$/" & c0 & "/^/" & c1 & "/^/" & c2 & "/^/" & c3 & "/^/" & c4 & "/^/" & c5 & "/^/" & c6 & "/^/" & c7 & "/^/" & c8 & "/^/" & c9 & "/^/" & c10 & "/^/" & c11 & "/^/" & c12 & "/^/" & c13 & "/^/"
End Function
Function GenerateInputParameter2(ParameterName As String, ParameterDatatype As String, ParameterSize As Long, ParameterValue As Variant) As Parameter
    Dim mvarParameterDatatype As DataTypeEnum
    Select Case ParameterDatatype
        Case "1"
            mvarParameterDatatype = adBigInt
        Case "2"
            mvarParameterDatatype = adBinary
        Case "3"
            mvarParameterDatatype = adBoolean
        Case "4"
            mvarParameterDatatype = adDouble
        Case "5"
            mvarParameterDatatype = adInteger
        Case "6"
            mvarParameterDatatype = adTinyInt
        Case "7"
            mvarParameterDatatype = adVarChar
        Case "8"
            mvarParameterDatatype = adVarWChar
        Case "9"
            mvarParameterDatatype = adVariant
    End Select
    GenerateInputParameter2.DataType = mvarParameterDatatype
    GenerateInputParameter2.Direction = adParamInput
    GenerateInputParameter2.Name = ParameterName
    GenerateInputParameter2.Size = ParameterSize
    GenerateInputParameter2.Value = ParameterValue
        
End Function

Public Function RunQuery2RecordSet(Query As String, Optional AdoConnection As ADODB.Connection) As ADODB.Recordset
    Dim MyAdoCommand As New ADODB.Command

    If AdoConnection Is Nothing Then Set AdoConnection = PosConnection
    
    If AdoConnection.ConnectionString = "" Then
        AdoConnection.ConnectionString = strConnectionString
    End If
    
    If AdoConnection.State = adStateClosed Then AdoConnection.Open strConnectionString
    AdoConnection.CursorLocation = adUseClient

'    With MyAdoCommand
'        .ActiveConnection = AdoConnection
'      '  .ActiveConnection.Source = Query
'        .CommandType = adCmdText
'        .CommandText = Query
'    End With
    Set RunQuery2RecordSet = New ADODB.Recordset
    With RunQuery2RecordSet
        .ActiveConnection = AdoConnection
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Source = Query
        .Open
    End With

'    Set RunQuery2RecordSet = MyAdoCommand.Execute
'    MyAdoCommand.Cancel

    Set MyAdoCommand = Nothing
End Function


