Attribute VB_Name = "KeysModule"
Option Explicit
Public objName As Object

 Public Sub KindKeys()
 If KindKey = 0 Then
    KindKey = 1
 End If
 End Sub
Public Sub CapsLock()
Dim i As Integer
If frmKeyBoard.FWKeyButton(KeyIndex).Index = 274 Then
For i = 200 To 225
    frmKeyBoard.FWKeyButton(i).Caption = Chr(i + 65 - 200)
    frmKeyBoard.FWKeyButton(i).Font.Name = "Times New Roman"
    frmKeyBoard.FWKeyButton(i).Font.Bold = True
    frmKeyBoard.FWKeyButton(i).Font.Size = 12
Next
   frmKeyBoard.FWKeyButton(226).Caption = ""
   frmKeyBoard.FWKeyButton(227).Caption = ""
   frmKeyBoard.FWKeyButton(228).Caption = ""
   frmKeyBoard.FWKeyButton(229).Caption = ""
   frmKeyBoard.FWKeyButton(230).Caption = ""
   frmKeyBoard.FWKeyButton(231).Caption = ""
   frmKeyBoard.FWKeyButton(232).Caption = ""
   frmKeyBoard.FWKeyButton(233).Caption = ""
   frmKeyBoard.FWKeyButton(276).Caption = ""
   frmKeyBoard.FWKeyButton(245).Caption = "?"
   frmKeyBoard.FWKeyButton(261).Caption = ","
   frmKeyBoard.FWKeyButton(253).Caption = ";"
End If
Select Case frmKeyBoard.FWKeyButton(KeyIndex).Index
       Case 263, 265, 274:
       Case 264:
            objName.Text = objName.Text + Space(1)
            objName.SetFocus
            SendKeys "{END}", True
       Case 266:
            SendKeys "{TAB}", False
            objName.SetFocus
       Case 269:
            SendKeys "{LEFT}", False
            objName.SetFocus
       Case 270:
            SendKeys "{RIGHT}", False
            objName.SetFocus
       Case 271:
            SendKeys "{UP}", False
            objName.SetFocus
       Case 272:
            SendKeys "{DOWN}", False
            objName.SetFocus
       Case 268:
            SendKeys "{BS}", False
            objName.SetFocus
       Case 273:
            SendKeys "{DEL}", False
            objName.SetFocus
       Case 275:
            SendKeys "{ENTER}", False
            objName.SetFocus
       Case Else:
            If TypeOf objName Is TextBox Then
                objName.RightToLeft = False
                objName.Alignment = 0
                objName.Text = objName.Text + (frmKeyBoard.FWKeyButton(KeyIndex).Caption)
                objName.SetFocus
                SendKeys "{end}", True
            End If
End Select

End Sub
Public Sub Latin()
Dim i As Integer
If frmKeyBoard.FWKeyButton(KeyIndex).Index = 263 Then
For i = 200 To 225
    frmKeyBoard.FWKeyButton(i).Caption = Chr(i + 97 - 200)
    frmKeyBoard.FWKeyButton(i).Font.Name = "Times New Roman"
    frmKeyBoard.FWKeyButton(i).Font.Bold = True
    frmKeyBoard.FWKeyButton(i).Font.Size = 12
Next
   frmKeyBoard.FWKeyButton(226).Caption = ""
   frmKeyBoard.FWKeyButton(227).Caption = ""
   frmKeyBoard.FWKeyButton(228).Caption = ""
   frmKeyBoard.FWKeyButton(229).Caption = ""
   frmKeyBoard.FWKeyButton(230).Caption = ""
   frmKeyBoard.FWKeyButton(231).Caption = ""
   frmKeyBoard.FWKeyButton(232).Caption = ""
   frmKeyBoard.FWKeyButton(233).Caption = ""
   frmKeyBoard.FWKeyButton(276).Caption = ""
   frmKeyBoard.FWKeyButton(245).Caption = "?"
   frmKeyBoard.FWKeyButton(261).Caption = ","
   frmKeyBoard.FWKeyButton(253).Caption = ";"
End If
Select Case frmKeyBoard.FWKeyButton(KeyIndex).Index
       Case 263, 265, 274:
       Case 264:
            objName.Text = objName.Text + Space(1)
            objName.SetFocus
            SendKeys "{END}", True
       Case 266:
            SendKeys "{TAB}", False
            objName.SetFocus
       Case 269:
            SendKeys "{LEFT}", False
            objName.SetFocus
       Case 270:
            SendKeys "{RIGHT}", False
            objName.SetFocus
       Case 271:
            SendKeys "{UP}", False
            objName.SetFocus
       Case 272:
            SendKeys "{DOWN}", False
            objName.SetFocus
       Case 268:
            SendKeys "{BS}", False
            objName.SetFocus
       Case 273:
            SendKeys "{DEL}", False
            objName.SetFocus
       Case 275:
            SendKeys "{ENTER}", False
            objName.SetFocus
       Case Else:
            If TypeOf objName Is TextBox Then
                objName.RightToLeft = False
                objName.Alignment = 0
                objName.Text = objName.Text + (frmKeyBoard.FWKeyButton(KeyIndex).Caption)
                objName.SetFocus
                SendKeys "{end}", True
            End If
End Select

End Sub

Public Sub Persian()
Dim i As Integer
If frmKeyBoard.FWKeyButton(KeyIndex).Index = 265 Then
For i = 200 To 231
    frmKeyBoard.FWKeyButton(i).Font.Name = "Traffic"
    frmKeyBoard.FWKeyButton(i).Font.Bold = True
    frmKeyBoard.FWKeyButton(i).Font.Size = 12
Next
   frmKeyBoard.FWKeyButton(215).Caption = "«"
   frmKeyBoard.FWKeyButton(214).Caption = "»"
   frmKeyBoard.FWKeyButton(213).Caption = "Å"
   frmKeyBoard.FWKeyButton(212).Caption = " "
   frmKeyBoard.FWKeyButton(211).Caption = "À"
   frmKeyBoard.FWKeyButton(210).Caption = "Ã"
   frmKeyBoard.FWKeyButton(209).Caption = "ç"
   frmKeyBoard.FWKeyButton(208).Caption = "Õ"
   frmKeyBoard.FWKeyButton(207).Caption = "Œ"
   frmKeyBoard.FWKeyButton(206).Caption = "œ"
   frmKeyBoard.FWKeyButton(205).Caption = "–"
   frmKeyBoard.FWKeyButton(204).Caption = "—"
   frmKeyBoard.FWKeyButton(203).Caption = "“"
   frmKeyBoard.FWKeyButton(202).Caption = "é"
   frmKeyBoard.FWKeyButton(201).Caption = "”"
   frmKeyBoard.FWKeyButton(200).Caption = "‘"
   frmKeyBoard.FWKeyButton(231).Caption = "’"
   frmKeyBoard.FWKeyButton(230).Caption = "÷"
   frmKeyBoard.FWKeyButton(229).Caption = "ÿ"
   frmKeyBoard.FWKeyButton(228).Caption = "Ÿ"
   frmKeyBoard.FWKeyButton(227).Caption = "⁄"
   frmKeyBoard.FWKeyButton(226).Caption = "€"
   frmKeyBoard.FWKeyButton(225).Caption = "›"
   frmKeyBoard.FWKeyButton(224).Caption = "ﬁ"
   frmKeyBoard.FWKeyButton(223).Caption = "ﬂ"
   frmKeyBoard.FWKeyButton(222).Caption = "ê"
   frmKeyBoard.FWKeyButton(221).Caption = "·"
   frmKeyBoard.FWKeyButton(220).Caption = "„"
   frmKeyBoard.FWKeyButton(219).Caption = "‰"
   frmKeyBoard.FWKeyButton(218).Caption = "Ê"
   frmKeyBoard.FWKeyButton(217).Caption = "Â"
   frmKeyBoard.FWKeyButton(216).Caption = "Ì"
   
   frmKeyBoard.FWKeyButton(245).Caption = "ø"
   frmKeyBoard.FWKeyButton(261).Caption = "°"
   frmKeyBoard.FWKeyButton(253).Caption = "∫"
   frmKeyBoard.FWKeyButton(232).Caption = "¡"
   frmKeyBoard.FWKeyButton(233).Caption = "¬"
   frmKeyBoard.FWKeyButton(276).Caption = "∆"
End If
Select Case frmKeyBoard.FWKeyButton(KeyIndex).Index
       Case 263, 265, 274:
       Case 264:
            objName.Text = objName.Text + Space(1)
            objName.SetFocus
            SendKeys "{END}", True
       Case 266:
            SendKeys "{TAB}", False
            objName.SetFocus
       Case 269:
            SendKeys "{LEFT}", False
            objName.SetFocus
       Case 270:
            SendKeys "{RIGHT}", False
            objName.SetFocus
       Case 271:
            SendKeys "{UP}", False
            objName.SetFocus
       Case 272:
            SendKeys "{DOWN}", False
            objName.SetFocus
       Case 268:
            SendKeys "{BS}", False
            objName.SetFocus
       Case 273:
            SendKeys "{DEL}", False
            objName.SetFocus
       Case 275:
            SendKeys "{ENTER}", False
            objName.SetFocus
       Case Else:
            If TypeOf objName Is TextBox Then
               ' objName.RightToLeft = True
               ' objName.Alignment = 1
                objName.Text = objName.Text + (frmKeyBoard.FWKeyButton(KeyIndex).Caption)
                objName.SetFocus
'                SendKeys "{end}", True
            End If
End Select
End Sub

Public Sub txtProperty(Obj As TextBox, Left As Integer, Top As Integer, Height As Integer, Width As Integer, FontSize As Integer)
       Obj.Left = Left
       Obj.Top = Top
       Obj.Height = Height
       Obj.Width = Width
       Obj.Font.Size = FontSize
End Sub
Public Sub lblProperty(Obj As Label, Left As Integer, Top As Integer, Height As Integer, Width As Integer, FontSize As Integer)
       Obj.Left = Left
       Obj.Top = Top
       Obj.Height = Height
       Obj.Width = Width
       Obj.Font.Size = FontSize
End Sub


        
