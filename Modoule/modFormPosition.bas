Attribute VB_Name = "modFormPosition"
Public Sub CenterCenter(ByRef MyForm As Form)
    MyForm.Left = (Screen.Width - MyForm.Width) / 2
    MyForm.Top = (Screen.Height - MyForm.Height) / 4
End Sub
Public Sub CenterTop(ByRef MyForm As Form)
    MyForm.Left = (Screen.Width - MyForm.Width) / 2
    MyForm.Top = 0
End Sub

Public Sub CenterBottom(ByRef MyForm As Form)
    MyForm.Left = (Screen.Width - MyForm.Width) / 2
    MyForm.Top = (Screen.Height - MyForm.Height)
End Sub

Public Sub LeftCenter(ByRef MyForm As Form)
    MyForm.Left = 0
    MyForm.Top = (Screen.Height - MyForm.Height) / 4
End Sub

Public Sub RightCenter(ByRef MyForm As Form)
    MyForm.Left = (Screen.Width - MyForm.Width)
    MyForm.Top = (Screen.Height - MyForm.Height) / 4
End Sub

Public Sub LeftTop(ByRef MyForm As Form)
    MyForm.Left = 0
    MyForm.Top = 0
End Sub

Public Sub RightTop(ByRef MyForm As Form)
    MyForm.Left = (Screen.Width - MyForm.Width)
    MyForm.Top = 0
End Sub

Public Sub LeftBottom(ByRef MyForm As Form)
    MyForm.Left = 0
    MyForm.Top = (Screen.Height - MyForm.Height)
End Sub

Public Sub RightButtom(ByRef MyForm As Form)
    MyForm.Left = (Screen.Width - MyForm.Width)
    MyForm.Top = (Screen.Height - MyForm.Height)
End Sub

Public Sub CenterCenterinSecondScreen(ByRef MyForm As Form)
     
    If MyForm.MDIChild = True Then
        MyForm.Left = mdifrm.Left + (mdifrm.Width - MyForm.Width) / 2
        MyForm.Top = (Screen.Height - MyForm.Height) / 4
    ElseIf LCase(strMainKey) = "total2" Then
        MyForm.Left = Screen.Width + (Screen.Width - MyForm.Width) / 2
        MyForm.Top = (Screen.Height - MyForm.Height) / 4
    ElseIf LCase(strMainKey) = "total" Then
        MyForm.Left = (Screen.Width - MyForm.Width) / 2
        MyForm.Top = (Screen.Height - MyForm.Height) / 4
    End If
End Sub
