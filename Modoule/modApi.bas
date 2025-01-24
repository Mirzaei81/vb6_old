Attribute VB_Name = "modApi"

'This fucntion changes the locale and as a result, the keyboardlayout gets adjusted

'parameters for api's
Public Const MF_BYPOSITION = &H400&
Const KL_NAMELENGTH As Long = 9                      'length of the keyboardbuffer
Const KLF_ACTIVATE  As Long = &H1                     'activate the layout

'the language constants
Public Const LANG_NL_STD As String = "00000413"
Public Const LANG_EN_US As String = "00000409"
Public Const LANG_DU_STD As String = "00000407"
Public Const LANG_FR_STD As String = "0000040C"
Public Const LANG_Pr_IR As String = "00000429"

'api's to adjust the keyboardlayout
Private Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Private Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal FLAGS As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOZORDER = 4
Public Const SWP_SHOWWINDOW = 40

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Declare Function SetParent Lib "user32" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
Declare Function SetWindowPos Lib "user32" _
      (ByVal hWnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal cX As Long, _
      ByVal cY As Long, _
      ByVal wFlags As Long) As Long
    
    Public Declare Function GetDesktopWindow Lib "user32" () As Long
      
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Long, ByVal bErase As Long) As Long
    Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
    Public Const GW_CHILD = 5
    Public Const WS_EX_LAYOUTRTL = &H400000
    Public Const GWL_EXSTYLE = (-20)
     


 
    Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
 
    If Topmost = True Then 'Makethewindowtopmost
        SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
        SetTopMostWindow = False
    End If
 
    End Function
Public Function SetKbLayout(strLocaleId As String) As Boolean
    'Changes the KeyboardLayout
    'Returns TRUE when the KeyboardLayout was adjusted properly, FALSE otherwise
    'If the KeyboardLayout isn't installed, this function will install it for you
    On Error Resume Next
    Dim strLocId As String 'used to retrieve current KeyboardLayout
    Dim strMsg As String   'used as buffer
    Dim lngErrNr As Long   'receives the API-error number

  'create a buffer
  strLocId = String(KL_NAMELENGTH, 0)
  'retrieve the current KeyboardLayout
  GetKeyboardLayoutName strLocId
  'Check whether the current KeyboardLayout and the
  'new one are the same
  If strLocId = (strLocaleId & Chr(0)) Then
    'If they're the same, we return immediately
    SetKbLayout = True
  Else
    'create buffer
    strLocId = String(KL_NAMELENGTH, 0)
    'load and activate the layout for the current thread
    strLocId = LoadKeyboardLayout((strLocaleId & Chr(0)), KLF_ACTIVATE)
    If IsNull(strLocId) Then  'returns NULL when it fails
      SetKbLayout = False
    Else 'check again
      'create buffer
      strLocId = String(KL_NAMELENGTH, 0)
      'retrieve the current layout
      GetKeyboardLayoutName strLocId
      If strLocId = (strLocaleId & Chr(0)) Then
        SetKbLayout = True
      Else
        SetKbLayout = False
      End If
    End If
  End If
End Function

Public Function GetKbLayout() As String
    On Error Resume Next
    Dim strLocId As String 'used to retrieve current KeyboardLayout
    
    'create a buffer
    strLocId = String(KL_NAMELENGTH, 0)
    
    'retrieve the current KeyboardLayout
    GetKeyboardLayoutName strLocId
    GetKbLayout = strLocId
End Function
