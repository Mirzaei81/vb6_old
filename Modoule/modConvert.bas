Attribute VB_Name = "modConvert"
Public Function ConvertToBin(intNum As Integer, intStringLen As String) As String
Const divisor = 2
Dim dividend  As Integer
Dim remainder As Integer

ConvertToBin = ""
dividend = intNum
Do
    If dividend < divisor Then
          ConvertToBin = CStr(dividend) & ConvertToBin
          Exit Do
    End If
    remainder = dividend Mod divisor
    dividend = dividend \ divisor
    ConvertToBin = CStr(remainder) & ConvertToBin
Loop
If Len(ConvertToBin) < intStringLen Then
    ConvertToBin = String(intStringLen - Len(ConvertToBin), Asc("0")) & ConvertToBin
End If
End Function

Public Function ConvertBinToInt(strBin As String) As Integer
Dim i As Integer
For i = 1 To Len(strBin)
    If Mid(strBin, i, 1) <> "0" And Mid(strBin, i, 1) <> "1" Then
        ConvertBinToInt = -1
        Exit Function
    End If
Next i
ConvertBinToInt = 0
For i = 1 To Len(strBin)
    ConvertBinToInt = ConvertBinToInt + (CInt(Mid(strBin, i, 1)) * (2 ^ (Len(strBin) - i)))
Next i
End Function

