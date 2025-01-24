VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicGoodMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_ViewSegmant(0 To 5) As Boolean
Private m_HeaderFont(0 To 5) As String
Private m_HeaderSizeFont(0 To 5) As String
Private m_HeaderColorFont(0 To 5) As String
Private m_GridFont(0 To 5) As String
Private m_GridSizeFont(0 To 5) As String
Private m_GridColorFont(0 To 5) As String
Private m_ViewRow(0 To 5) As Boolean
Private m_ViewName(0 To 5) As Boolean
Private m_ViewFee1(0 To 5) As Boolean
Private m_ViewFee2(0 To 5) As Boolean
Private m_ViewPicture(0 To 5) As Boolean
Private m_ViewDescription(0 To 5) As Boolean
Private m_RowName(0 To 5) As String
Private m_GoodName(0 To 5) As String
Private m_Fee1Name(0 To 5) As String
Private m_Fee2Name(0 To 5) As String
Private m_HeaderTitr(0 To 5) As String

Public Property Get ViewSegmant(Index) As Boolean
    ViewSegmant = m_ViewSegmant(Index)
End Property
Public Property Let ViewSegmant(Index, ByVal vData As Boolean)
    m_ViewSegmant(Index) = vData
End Property
Public Property Get HeaderFont(Index) As String
    HeaderFont = m_HeaderFont(Index)
End Property
Public Property Let HeaderFont(Index, ByVal vData As String)
    m_HeaderFont(Index) = vData
End Property
Public Property Get HeaderSizeFont(Index) As String
    HeaderSizeFont = m_HeaderSizeFont(Index)
End Property
Public Property Let HeaderSizeFont(Index, ByVal vData As String)
    m_HeaderSizeFont(Index) = vData
End Property
Public Property Get HeaderColorFont(Index) As String
    HeaderColorFont = m_HeaderColorFont(Index)
End Property
Public Property Let HeaderColorFont(Index, ByVal vData As String)
    m_HeaderColorFont(Index) = vData
End Property
Public Property Get GridFont(Index) As String
    GridFont = m_GridFont(Index)
End Property
Public Property Let GridFont(Index, ByVal vData As String)
    m_GridFont(Index) = vData
End Property
Public Property Get GridSizeFont(Index) As String
    GridSizeFont = m_GridSizeFont(Index)
End Property
Public Property Let GridSizeFont(Index, ByVal vData As String)
    m_GridSizeFont(Index) = vData
End Property
Public Property Get GridColorFont(Index) As String
    GridColorFont = m_GridColorFont(Index)
End Property
Public Property Let GridColorFont(Index, ByVal vData As String)
    m_GridColorFont(Index) = vData
End Property
Public Property Get ViewRow(Index) As Boolean
    ViewRow = m_ViewRow(Index)
End Property
Public Property Let ViewRow(Index, ByVal vData As Boolean)
    m_ViewRow(Index) = vData
End Property
Public Property Get ViewName(Index) As Boolean
    ViewName = m_ViewName(Index)
End Property
Public Property Let ViewName(Index, ByVal vData As Boolean)
    m_ViewName(Index) = vData
End Property
Public Property Get ViewFee1(Index) As Boolean
    ViewFee1 = m_ViewFee1(Index)
End Property
Public Property Let ViewFee1(Index, ByVal vData As Boolean)
    m_ViewFee1(Index) = vData
End Property
Public Property Get ViewFee2(Index) As Boolean
    ViewFee2 = m_ViewFee2(Index)
End Property
Public Property Let ViewFee2(Index, ByVal vData As Boolean)
    m_ViewFee2(Index) = vData
End Property
Public Property Get ViewPicture(Index) As Boolean
    ViewPicture = m_ViewPicture(Index)
End Property
Public Property Let ViewPicture(Index, ByVal vData As Boolean)
    m_ViewPicture(Index) = vData
End Property
Public Property Get ViewDescription(Index) As Boolean
    ViewDescription = m_ViewDescription(Index)
End Property
Public Property Let ViewDescription(Index, ByVal vData As Boolean)
    m_ViewDescription(Index) = vData
End Property
Public Property Get RowName(Index) As String
    RowName = m_RowName(Index)
End Property
Public Property Let RowName(Index, ByVal vData As String)
    m_RowName(Index) = vData
End Property
Public Property Get GoodName(Index) As String
    GoodName = m_GoodName(Index)
End Property
Public Property Let GoodName(Index, ByVal vData As String)
    m_GoodName(Index) = vData
End Property
Public Property Get Fee1Name(Index) As String
    Fee1Name = m_Fee1Name(Index)
End Property
Public Property Let Fee1Name(Index, ByVal vData As String)
    m_Fee1Name(Index) = vData
End Property
Public Property Get Fee2Name(Index) As String
    Fee2Name = m_Fee2Name(Index)
End Property
Public Property Let Fee2Name(Index, ByVal vData As String)
    m_Fee2Name(Index) = vData
End Property
Public Property Get HeaderTitr(Index) As String
    HeaderTitr = m_HeaderTitr(Index)
End Property
Public Property Let HeaderTitr(Index, ByVal vData As String)
    m_HeaderTitr(Index) = vData
End Property

Private Sub Class_Initialize()
    
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim str As String
    Dim LenghStr As Integer
    
    Dim IsFileExist As Boolean
    Dim i As Long
    If GoodMenuSettingFile = "" Then End    'Only  For  Make Exe File
    
    For i = 0 To 5
        Set tempstring = filetemp.OpenTextFile(GoodMenuSettingFile, ForReading, False, TristateFalse)
        Do While tempstring.AtEndOfLine = False
            str = tempstring.ReadLine
            LenghStr = InStr(1, str, "=", vbTextCompare)
            
            If InStr(1, str, "ViewSegmant" & i, vbTextCompare) Then
               m_ViewSegmant(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "HeaderTitr" & i, vbTextCompare) Then
               m_HeaderTitr(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "HeaderFont" & i, vbTextCompare) Then
               m_HeaderFont(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "HeaderSizeFont" & i, vbTextCompare) Then
               m_HeaderSizeFont(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "HeaderColorFont" & i, vbTextCompare) Then
               m_HeaderColorFont(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "GridFont" & i, vbTextCompare) Then
               m_GridFont(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "GridSizeFont" & i, vbTextCompare) Then
               m_GridSizeFont(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "GridColorFont" & i, vbTextCompare) Then
               m_GridColorFont(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "ViewRow" & i, vbTextCompare) Then
               m_ViewRow(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "ViewName" & i, vbTextCompare) Then
               m_ViewName(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "ViewFee1" & i, vbTextCompare) Then
               m_ViewFee1(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "ViewFee2" & i, vbTextCompare) Then
               m_ViewFee2(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "ViewPicture" & i, vbTextCompare) Then
               m_ViewPicture(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "ViewDescription" & i, vbTextCompare) Then
               m_ViewDescription(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "RowName" & i, vbTextCompare) Then
               m_RowName(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "GoodName" & i, vbTextCompare) Then
               m_GoodName(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "Fee1Name" & i, vbTextCompare) Then
               m_Fee1Name(i) = Mid(str, LenghStr + 1)
            
            ElseIf InStr(1, str, "Fee2Name" & i, vbTextCompare) Then
               m_Fee2Name(i) = Mid(str, LenghStr + 1)
            
            End If
        Loop
     tempstring.Close
     Next i

End Sub






