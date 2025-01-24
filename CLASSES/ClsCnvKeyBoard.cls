VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCnvKeyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarConvert As Integer 'local copy

Public Property Let Convert(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Convert = 5
    mvarConvert = vData
End Property

Public Property Get Convert() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Convert
    Convert = Me.CnvKeyBoard(mvarConvert)
End Property

Public Function CnvKeyBoard(mydata As Integer) As Integer

Select Case clsStation.WinAscii

      Case 1
         Select Case mydata
         
            Case 129:
                CnvKeyBoard = 92
            Case 141:
                CnvKeyBoard = 93
            Case 142:
                CnvKeyBoard = 67
            Case 144:
                CnvKeyBoard = 39
            Case 152:
                 CnvKeyBoard = 59
            Case 155:
                CnvKeyBoard = 67
            Case 161:
                CnvKeyBoard = 84
            Case 171:
                CnvKeyBoard = 75
             Case 186:
                CnvKeyBoard = 89
            Case 187:
                CnvKeyBoard = 76
            Case 191:
                CnvKeyBoard = 63
            Case 193:
                CnvKeyBoard = 77
            Case 194:
                CnvKeyBoard = 72
            Case 195:
                CnvKeyBoard = 78
            Case 196:
                CnvKeyBoard = 86
            Case 197:
                CnvKeyBoard = 66
            Case 198:
                CnvKeyBoard = 109
            Case 199:
                CnvKeyBoard = 104
            Case 200:
                CnvKeyBoard = 102
            Case 201:
                CnvKeyBoard = 90
            Case 202:
                CnvKeyBoard = 106
            Case 203:
                CnvKeyBoard = 101
            Case 204:
                CnvKeyBoard = 91
            Case 205:
                CnvKeyBoard = 112
            Case 206:
                CnvKeyBoard = 111
            Case 207:
                CnvKeyBoard = 110
            Case 208:
                CnvKeyBoard = 98
        '    Case 209:
        '        CnvKeyBoard = 82
            Case 209:
                CnvKeyBoard = 118
            Case 210:
                CnvKeyBoard = 99
            Case 211:
                CnvKeyBoard = 115
            Case 212:
                CnvKeyBoard = 97
            Case 213:
                CnvKeyBoard = 119
            Case 214:
                CnvKeyBoard = 113
            Case 215:
                CnvKeyBoard = 126
            Case 216:
                CnvKeyBoard = 120
            Case 217:
                CnvKeyBoard = 122
            Case 218:
                CnvKeyBoard = 117
            Case 219:
                CnvKeyBoard = 121
            Case 220:
                CnvKeyBoard = 74
            Case 221:
                CnvKeyBoard = 116
            Case 222:
                CnvKeyBoard = 114
            Case 225:
                CnvKeyBoard = 103
            Case 227:
                CnvKeyBoard = 108
            Case 228:
                CnvKeyBoard = 107
            Case 229:
                CnvKeyBoard = 105
            Case 230:
                CnvKeyBoard = 44
        '    Case 237:
        '        CnvKeyBoard = 88
            Case 237:
                CnvKeyBoard = 100
            Case 240:
                CnvKeyBoard = 81
            Case 241:
                CnvKeyBoard = 87
            Case 242:
                CnvKeyBoard = 69
            Case 243:
                CnvKeyBoard = 65
            Case 245:
                CnvKeyBoard = 83
            Case 246:
                CnvKeyBoard = 68
            Case 247:
                CnvKeyBoard = 96
            Case 248:
                CnvKeyBoard = 70
            Case 40
                If Val(GetKbLayout) = Val(LANG_Pr_IR) Then
                    CnvKeyBoard = 41
                Else
                    CnvKeyBoard = 40
                End If
            Case 41
                If Val(GetKbLayout) = Val(LANG_Pr_IR) Then
                    CnvKeyBoard = 40
                Else
                    CnvKeyBoard = 41
                End If
            Case 123
                If Val(GetKbLayout) = Val(LANG_Pr_IR) Then
                    CnvKeyBoard = 125
                Else
                    CnvKeyBoard = 123
                End If
            Case 125
                If Val(GetKbLayout) = Val(LANG_Pr_IR) Then
                    CnvKeyBoard = 123
                Else
                    CnvKeyBoard = 125
                End If
            Case Else
                CnvKeyBoard = mydata
           
           End Select
            
         Case 2
          
          Select Case mydata
         
            Case 46:
               CnvKeyBoard = 62
            Case 129:
                CnvKeyBoard = 92
            Case 141:
                CnvKeyBoard = 93
            Case 142:
                CnvKeyBoard = 67
            Case 144:
                CnvKeyBoard = 39
            Case 152:
                 CnvKeyBoard = 59
            Case 155:
                CnvKeyBoard = 74
            Case 157:
                CnvKeyBoard = 125
            Case 158:
                CnvKeyBoard = 123
             Case 139:
                CnvKeyBoard = 75
            Case 161:
                CnvKeyBoard = 84
            Case 171:
                CnvKeyBoard = 86
             Case 186:
                CnvKeyBoard = 89
            Case 187:
                CnvKeyBoard = 76
            Case 191:
                CnvKeyBoard = 63
            Case 193:
                CnvKeyBoard = 77
            Case 194:
                CnvKeyBoard = 72
            Case 195:
                CnvKeyBoard = 78
            Case 196:
                CnvKeyBoard = 86
            Case 197:
                CnvKeyBoard = 66
            Case 198:
                CnvKeyBoard = 109
            Case 199:
                CnvKeyBoard = 104
            Case 200:
                CnvKeyBoard = 102
            Case 201:
                CnvKeyBoard = 90
            Case 202:
                CnvKeyBoard = 106
            Case 203:
                CnvKeyBoard = 101
            Case 204:
                CnvKeyBoard = 91
            Case 205:
                CnvKeyBoard = 112
            Case 206:
                CnvKeyBoard = 111
            Case 207:
                CnvKeyBoard = 110
            Case 208:
                CnvKeyBoard = 98
        '    Case 209:
        '        CnvKeyBoard = 82
            Case 209:
                CnvKeyBoard = 118
            Case 210:
                CnvKeyBoard = 99
            Case 211:
                CnvKeyBoard = 115
            Case 212:
                CnvKeyBoard = 97
            Case 213:
                CnvKeyBoard = 119
            Case 214:
                CnvKeyBoard = 113
            Case 215:
                CnvKeyBoard = 126
            Case 216:
                CnvKeyBoard = 120
            Case 217:
                CnvKeyBoard = 122
            Case 218:
                CnvKeyBoard = 117
            Case 219:
                CnvKeyBoard = 121
            Case 220:
                CnvKeyBoard = 74
            Case 221:
                CnvKeyBoard = 116
            Case 222:
                CnvKeyBoard = 114
            Case 223:
                CnvKeyBoard = 59
            Case 225:
                CnvKeyBoard = 103
            Case 227:
                CnvKeyBoard = 108
            Case 228:
                CnvKeyBoard = 107
            Case 229:
                CnvKeyBoard = 105
            Case 230:
                CnvKeyBoard = 44
        '    Case 237:
        '        CnvKeyBoard = 88
            Case 237:
                CnvKeyBoard = 100
            Case 240:
                CnvKeyBoard = 81
            Case 241:
                CnvKeyBoard = 87
            Case 242:
                CnvKeyBoard = 69
            Case 243:
                CnvKeyBoard = 65
            Case 245:
                CnvKeyBoard = 83
            Case 246:
                CnvKeyBoard = 68
            Case 247:
                CnvKeyBoard = 96
            Case 248:
                CnvKeyBoard = 70
            Case 250:
                CnvKeyBoard = 71
            Case 40
                If Val(GetKbLayout) = Val(LANG_Pr_IR) Then
                    CnvKeyBoard = 41
                Else
                    CnvKeyBoard = 40
                End If
            Case 41
                If Val(GetKbLayout) = Val(LANG_Pr_IR) Then
                    CnvKeyBoard = 40
                Else
                    CnvKeyBoard = 41
                End If
            Case 123
                If Val(GetKbLayout) = Val(LANG_Pr_IR) Then
                    CnvKeyBoard = 125
                Else
                    CnvKeyBoard = 123
                End If
            Case 125
                If Val(GetKbLayout) = Val(LANG_Pr_IR) Then
                    CnvKeyBoard = 123
                Else
                    CnvKeyBoard = 125
                End If
            Case Else
                CnvKeyBoard = mydata
         
         End Select
End Select
'If Not (CnvKeyBoard > 0) Then
'    CnvKeyBoard = 0
'End If
End Function


