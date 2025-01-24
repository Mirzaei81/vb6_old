VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{687FF23F-0E9B-449D-A782-2C5E7116EDB0}#1.0#0"; "Fardate.ocx"
Begin VB.Form frmRep 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin FarDate1.FarDate FarDate2 
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   670
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Object.ToolTipText     =   "FarDate"
   End
   Begin FarDate1.FarDate FarDate1 
      Height          =   350
      Left            =   3000
      TabIndex        =   1
      Top             =   670
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Object.ToolTipText     =   "FarDate"
   End
   Begin Total.UCReportIO UCReportIO1 
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   10081
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2160
      Top             =   1590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "frmRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private T1 As String
Private T2 As String
Private T3 As String
Private TDate As String
Private DateBefor As String
Private DateAfter As String
Private Time1 As String
Private Time2 As String
Private clsDate As New clsDate
Private iHeight As Integer
Private iWidth As Integer
Private RepSellerType As Integer
Private cust1, cust2 As Long
Private User1, User2, Shift1, Shift2, Station1, Station2, PrintFormat, Sup1, Sup2 As Integer
Private GoodType1, GoodType2 As Double
Private StoreType1, StoreType2, Inventory1, Inventory2, GoodFlag, InventoryType, CustomerType, DetailType, PaperType, OrderKind, ShowKind As Integer
Private GoodLevel1FromCode, GoodLevel1ToCode As Double
Private GoodLevel2FromCode, GoodLevel2ToCode As Double
Private FactorNo1, FactorNo2 As Long
Private SPrice1, SPrice2 As Double
Private Discont1, Discont2 As Double
Dim Parameter() As Parameter
'Dim ArrayUbound As Integer
Dim ReportSP_Name As String
Dim Statusvar As Integer
Dim ReportHeader As String
Dim CalculatedRialType, SortOrder As Integer
Dim b1, b2 As Integer
Dim MonthBefore, MonthAfter As Integer
Dim FromGCodeL1, toGCodeL1 As Integer
Dim AccountYear1, AccountYear2 As Integer
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Private Sub FarDate1_Change()
    Me.UCReportIO1.txt(0) = Mid(FarDate1.Text, 3, 8)
End Sub


Private Sub FarDate2_Change()
    Me.UCReportIO1.txt(1) = Mid(FarDate2.Text, 3, 8)
End Sub

Private Sub Form_Activate()

    VarActForm = Me.Name

    Me.UCReportIO1.ArrangeCmd = Me.UCReportIO1.ConditionNo
    iHeight = UCReportIO1.Height + 10
    iWidth = UCReportIO1.Width - 100
    Me.Height = iHeight
    Me.Width = iWidth
    modgl.RightButton False
End Sub

Private Sub Form_Deactivate()
'        Call RemoveIndex
        Me.UCReportIO1.RemoveIndex
End Sub

Private Sub Form_Load()

''''    If ClsFormAccess.frmRep = False Then
''''        Unload Me
''''        Exit Sub
''''    End If


    Dim strCnn As String
    Dim i As Integer
    
    RightTop Me
    Me.Left = Me.Left - 1000
    
    VarActForm = Me.Name
    
   
Select Case CRepFlag
    Case "RepGroupGoodsSale"
            Me.UCReportIO1.MSG = "ê“«—‘ ›—Ê‘ ê—ÊÂÌ ò«·« Â«"
            ReportHeader = "ê“«—‘ ›—Ê‘ ê—ÊÂÌ ò«·« Â«"
            Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
            Me.UCReportIO1.Add vbtextbox, 2, "Supplier", "Supplier", "1", "99999"
            
            Me.UCReportIO1.Add vbcombobox, 2, "tFacm.StationId", "StationId", "1", "999", "StationId", "StationId", "SELECT * FROM tStations  Order By StationId "
            Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
            Me.UCReportIO1.Add vbcombobox, 2, "level2", "L2", "1010", "9999", "Code", "Description", "SELECT * FROM tGoodLevel2 Order By Code "
            Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
            Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
          '  Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
            Me.UCReportIO1.Add vbcombobox, 2, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
    Case "RepGroupGoodsBuy"
            Me.UCReportIO1.MSG = "ê“«—‘ Œ—Ìœ ê—ÊÂÌ ﬂ«·«Â«"
            ReportHeader = "ê“«—‘ Œ—Ìœ ê—ÊÂÌ ﬂ«·«Â«"
            Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
            Me.UCReportIO1.Add vbtextbox, 2, "Supplier", "Supplier", "1", "99999"
            Me.UCReportIO1.Add vbcombobox, 2, "tFacm.StationId", "StationId", "1", "999", "StationID", "StationID", "SELECT * FROM tStations Order By StationId "
            Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
            Me.UCReportIO1.Add vbcombobox, 2, "level2", "L2", "1010", "9999", "Code", "Description", "SELECT * FROM tGoodLevel2 Order By Code "
            Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
            Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
            Me.UCReportIO1.Add vbcombobox, 2, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
'-----------------------------
    Case "RepFichSale"
            Me.UCReportIO1.MSG = "ê“«—‘ ›—Ê‘ »— «”«” ›Ì‘"
            ReportHeader = "ê“«—‘ ›—Ê‘ »— «”«” ›Ì‘"
            Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
            Me.UCReportIO1.Add vbcombobox, 2, "tFacm.[User]", "User", "1", "9999", "ppno", "nvcSurName", "SELECT * FROM tPer INNER JOIN tUser ON tPer.ppno = tUser.PPNo  where ActDeAct=1 Order By tPer.ppno"
            Me.UCReportIO1.Add vbcombobox, 2, "tFacm.StationId", "StationId", "1", "999", "StationID", "StationID", "SELECT * FROM tStations  Order By StationId "
            Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
            Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
    Case "RepFichBuy"
        
            Me.UCReportIO1.MSG = "ê“«—‘ Œ—Ìœ »— «”«” ›Ì‘"
            ReportHeader = "ê“«—‘ Œ—Ìœ »— «”«” ›Ì‘"
            Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
            Me.UCReportIO1.Add vbcombobox, 2, "tFacm.[User]", "User", "1", "9999", "ppno", "nvcSurName", "SELECT * FROM tPer INNER JOIN tUser ON tPer.ppno = tUser.PPNo where ActDeAct=1 Order By tPer.ppno"
            Me.UCReportIO1.Add vbcombobox, 2, "tFacm.StationId", "StationId", "1", "999", "StationID", "StationID", "SELECT * FROM tStations  Order By StationId "
            Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
            Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        
'-----------------------------
    Case "RepTimePerSell"
            Me.UCReportIO1.MSG = "¬„«— ›—Ê‘ ”«⁄ Ì Ê œ—’œÌ"
            ReportHeader = "¬„«— ›—Ê‘ ”«⁄ Ì Ê œ—’œÌ"
            Me.UCReportIO1.Add vbtextbox, 2, "[Date]", "Dat", "70/01/01", "99/01/01"
            Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
            Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
            Me.UCReportIO1.Add vbcombobox, 2, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft

'-----------------------------
    Case "RepCashInvoice"
            Me.UCReportIO1.MSG = "ê“«—‘ ›—Ê‘ ’‰œÊﬁ"
            ReportHeader = "ê“«—‘ ›—Ê‘ ’‰œÊﬁ"
            Me.UCReportIO1.Add vbtextbox, 2, "[Date]", "Dat", "70/01/01", "99/01/01"
            Me.UCReportIO1.Add vbcombobox, 2, "tFacm.[User]", "User", "1", "9999", "ppno", "nvcSurName", "SELECT * FROM tPer INNER JOIN tUser ON tPer.ppno = tUser.PPNo where ActDeAct=1 Order By tPer.ppno"
            Me.UCReportIO1.Add vbcombobox, 2, "tFacm.StationId", "StationId", "1", "999", "StationID", "StationID", "SELECT * FROM tStations  Order By StationId "
            Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
            Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "

'-----------------------
    Case "RepBedeh"
            Me.UCReportIO1.MSG = "ê“«—‘ »œÂﬂ«—Ì ÅÌﬂÂ«"
            ReportHeader = "ê“«—‘ »œÂﬂ«—Ì ÅÌﬂÂ«"
            Me.UCReportIO1.Add vbtextbox, 2, "[Date]", "Dat", "70/01/01", "99/01/01"
            Me.UCReportIO1.Add vbcombobox, 2, "tFacm.[InCharge]", "carrier", "1", "9999", "pPNO", "nvcSurName", "SELECT * FROM tPer WHERE job = " & EnumIncharge.Payk & " And  ActDeAct=1 Order By ppno"
            Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
            Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "

'-----------------------------
    Case "RepBedehDetail"
            Me.UCReportIO1.MSG = "ê“«—‘ »œÂﬂ«—Ì ÅÌﬂÂ« ‹ Ã“∆Ì« "
            ReportHeader = "ê“«—‘ »œÂﬂ«—Ì ÅÌﬂÂ« ‹ Ã“∆Ì« "
            Me.UCReportIO1.Add vbtextbox, 2, "[Date]", "Dat", "70/01/01", "99/01/01"
            Me.UCReportIO1.Add vbcombobox, 2, "tFacm.[InCharge]", "carrier", "1", "9999", "pPNO", "nvcSurName", "SELECT * FROM tPer WHERE job = " & EnumIncharge.Payk & " And   ActDeAct= 1Order By ppno"
            Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
            Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
       

    Case "RepGoodDifferences"
            Me.UCReportIO1.MSG = "ê“«—‘ ·Ì”   €ÌÌ—«  ò«·« Â«"
            ReportHeader = "ê“«—‘ ·Ì”   €ÌÌ—«  ò«·« Â«"
            Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
            Me.UCReportIO1.Add vbcombobox, 2, "level2", "L2", "1010", "9999", "Code", "Description", "SELECT * FROM tGoodLevel2 Order By Code "
            Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
    
    Case "RepGoodList"
            Me.UCReportIO1.MSG = "ê“«—‘ ·Ì”  ò«·« Â«Ì  ⁄—Ì› ‘œÂ"
            ReportHeader = "ê“«—‘ ·Ì”  ò«·« Â«Ì  ⁄—Ì› ‘œÂ"
            Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
            Me.UCReportIO1.Add vbcombobox, 2, "level2", "L2", "1010", "9999", "Code", "Description", "SELECT * FROM tGoodLevel2 Order By Code "
            Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
    
'-----------------------------
    Case "RepDetailGoodsSale", "RepDetailGoodsSaleReturn", "RepDetailGoodsBuy", "RepDetailGoodsBuyReturn"     '2 , 28 , 13 ,29
            If CRepFlag = "RepDetailGoodsSale" Then
                If clsStation.Language = Farsi Then
                    Me.UCReportIO1.MSG = "ê“«—‘ ›—Ê‘ —Ì“ ﬂ«·« Â«"
                    ReportHeader = "ê“«—‘ ›—Ê‘ —Ì“ ﬂ«·« Â«"
                Else
                    Me.UCReportIO1.MSG = "Detail Goods Sale Report"
                    ReportHeader = "Detail Goods Sale Report"
                End If
               Statusvar = 2
            ElseIf CRepFlag = "RepDetailGoodsSaleReturn" Then
                If clsStation.Language = Farsi Then
                    Me.UCReportIO1.MSG = "ê“«—‘ »—ê‘  «“ ›—Ê‘ —Ì“ ﬂ«·« Â«"
                    ReportHeader = "ê“«—‘ »—ê‘  «“ ›—Ê‘ —Ì“ ﬂ«·« Â«"
                Else
                    Me.UCReportIO1.MSG = "Detail Goods Sale Revocation Report"
                    ReportHeader = "Detail Goods Sale Revocation Report"
                End If
               Statusvar = 5
            ElseIf CRepFlag = "RepDetailGoodsBuy" Then
                If clsStation.Language = Farsi Then
                    Me.UCReportIO1.MSG = "ê“«—‘ Œ—Ìœ —Ì“ ﬂ«·« Â« "
                    ReportHeader = "ê“«—‘ Œ—Ìœ —Ì“ ﬂ«·« Â« "
                Else
                    Me.UCReportIO1.MSG = "Detail Goods Purchase Report "
                    ReportHeader = "Detail Goods Purchase Report "
                End If
               Statusvar = 1
            ElseIf CRepFlag = "RepDetailGoodsBuyReturn" Then
                If clsStation.Language = Farsi Then
                    Me.UCReportIO1.MSG = "ê“«—‘ »—ê‘  «“ Œ—Ìœ —Ì“ ﬂ«·« Â« "
                    ReportHeader = "ê“«—‘ »—ê‘  «“ Œ—Ìœ —Ì“ ﬂ«·« Â«"
                Else
                    Me.UCReportIO1.MSG = "Detail Goods Purchase Revocation Report "
                    ReportHeader = "Detail Goods Purchase Revocation Report "
                End If
               Statusvar = 4
            End If
            Me.UCReportIO1.Add vbtextbox, 2, "[Date]", "Dat", "70/01/01", "99/01/01"
            Me.UCReportIO1.Add vbtextbox, 2, "Supplier", "Supplier", "1", "99999"
            
            Me.UCReportIO1.Add vbcombobox, 2, "tFacm.StationId", "StationId", "1", "999", "StationID", "StationID", "SELECT * FROM tStations  Order By StationId "
            Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
            Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
            Me.UCReportIO1.Add vbcombobox, 2, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
   
   Case "RepServeKindSell"   '6
        Me.UCReportIO1.MSG = "ê“«—‘ «‰Ê«⁄ ›—Ê‘ "
        ReportHeader = "ê“«—‘ «‰Ê«⁄ ›—Ê‘ "
        Me.UCReportIO1.Add vbtextbox, 2, "[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 2, "tFacm.[User]", "User", "1", "9999", "ppno", "nvcSurName", "SELECT * FROM tPer INNER JOIN tUser ON tPer.ppno = tUser.PPNo  where ActDeAct=1 Order By tPer.ppno"
        Me.UCReportIO1.Add vbcombobox, 2, "tFacm.StationId", "StationId", "1", "999", "StationID", "StationID", "SELECT * FROM tStations  Order By StationId "
        Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
'-----------------------------
    Case "RepCustPricDiscount", "RepCustPricDiscountReturn"
        If CRepFlag = "RepCustPricDiscount" Then
           Me.UCReportIO1.MSG = "·Ì”  ›—Ê‘ »Â „‘ —ﬂÌ‰ Ê Œ—Ìœ«“ «„Ì‰ ﬂ‰‰œê«‰"
           Statusvar = 0
        ElseIf CRepFlag = "RepCustPricDiscountReturn" Then
           Me.UCReportIO1.MSG = " »—ê‘  «“ ›—Ê‘(„‘ —ﬂÌ‰)Ê Œ—Ìœ( «„Ì‰ ﬂ‰‰œê«‰)"
           Statusvar = 1
        End If
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbtextbox, 2, "Customer", "Customer", "1", "99999"
        Me.UCReportIO1.Add vbtextbox, 2, "sumprice", "SPrice", "0", "99999999999"
        Me.UCReportIO1.Add vbtextbox, 2, "DiscountTotal", "SDiscount", "0", "99999999999"
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "CustomerType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Cust Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
    '-----------------------------
    Case "RepCustPrice", "RepCustPriceReturn"
        If CRepFlag = "RepCustPrice" Then
           Me.UCReportIO1.MSG = "·Ì”  ›—Ê‘ »Â „‘ —ﬂÌ‰ ÊŒ—Ìœ «“  «„Ì‰ ﬂ‰‰œê«‰ - Ã“∆Ì« "
           Statusvar = 0
        ElseIf CRepFlag = "RepCustPriceReturn" Then
           Me.UCReportIO1.MSG = " »—ê‘  «“›—Ê‘(„‘ —ﬂÌ‰)ÊŒ—Ìœ( «„Ì‰ ﬂ‰‰œê«‰)- Ã“∆Ì« "
           Statusvar = 1
        End If
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbtextbox, 2, "Customer", "Customer", "1", "999999"
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "CustomerType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Cust Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "Details", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Goods Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
    '-----------------------------
''    Case "RepCustInfo"
''        Me.UCReportIO1.Msg = "ê“«—‘ «‰ Œ«» „‘ —ﬂÌ‰"
''        Me.UCReportIO1.Add vbtextbox, 2, "[Date]", "Dat", "70/01/01", "99/01/01"
''        Me.UCReportIO1.Add vbtextbox, 2, "tcust.Code", "Customer", "1", "99999"
    '-----------------------------
    Case "RepSystemGroup"
        Me.UCReportIO1.MSG = "ê“«—‘ ê—ÊÂÂ«Ì ”Ì” „Ì"
        ReportHeader = "ê“«—‘ ê—ÊÂÂ«Ì ”Ì” „Ì"
        Me.UCReportIO1.Add vbtextbox, 2, "[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        Me.UCReportIO1.Add vbcombobox, 2, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
    
    '-----------------------------
    Case "RepSerialFich"
        Me.UCReportIO1.MSG = "ê“«—‘ »— «”«” ‘„«—Â ”—Ì«·"
        ReportHeader = "ê“«—‘ »— «”«” ‘„«—Â ”—Ì«·"
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbtextbox, 2, "No", "SerialNo", "0", "99999999999"
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
    '-----------------------------
    Case "RepShift"
        Me.UCReportIO1.MSG = "ê“«—‘ ‘Ì› "
        ReportHeader = "ê“«—‘ ‘Ì› "
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 2, "tFacm.[User]", "User", "1", "9999", "ppno", "nvcSurName", "SELECT * FROM tPer INNER JOIN tUser ON tPer.ppno = tUser.PPNo  where ActDeAct=1 Order By tPer.ppno"
        Me.UCReportIO1.Add vbcombobox, 2, "tfacD.ShiftNo", "Shift", "0", "9", "Code", "Description", "SELECT * FROM tshift  order by code"
        Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
    '-----------------------------
    Case "RepMojodi"
       Me.UCReportIO1.MSG = "ê“«—‘ „ÊÃÊœÌ  ⁄œ«œÌ «‰»«— - œ—  «—ÌŒ Œ«’"
        Me.UCReportIO1.Add vbtextbox, 1, "tFacM.[Date]", "Dat", "70/01/01" '', "99/01/01"
       ''' Me.UCReportIO1.Add vbtextbox, 2, "Supplier", "Supplier", "1", "99999"
        Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
        Me.UCReportIO1.Add vbcombobox, 2, "InventoryNo", "Description", "1", "99", "InventoryNo", "Description", "SELECT * FROM tInventory Order By InventoryNo "
''''        Me.UCReportIO1.Add vbcombobox, 1, "Code", "StoreDescription", "0", "9", "Code", "Description", "SELECT * FROM tblpub_StoreKind Order By Code "
       '''' Me.UCReportIO1.Add vbcombobox, 1, "Code", "CalculatedGood", "0", "1", "Code", "Description", "SELECT * FROM tblpub_CalculatedGood Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "CalculatedGood", "0", "1", "Code", "Description", "SELECT * FROM tblpub_CalculatedGood Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "SortOrder", "0", "1", "Code", "Description", "SELECT * FROM tblpub_Sort Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
    
    Case "RepMojodiRial"
         Me.UCReportIO1.MSG = "ê“«—‘ „ÊÃÊœÌ —Ì«·Ì «‰»«— - œ—  «—ÌŒ Œ«’ "
        Me.UCReportIO1.Add vbtextbox, 1, "tFacM.[Date]", "Dat", "70/01/01" ''', "99/01/01"
'''        Me.UCReportIO1.Add vbtextbox, 2, "Supplier", "Supplier", "1", "99999"
        Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
        Me.UCReportIO1.Add vbcombobox, 2, "InventoryNo", "Description", "1", "99", "InventoryNo", "Description", "SELECT * FROM tInventory Order By InventoryNo "
''''        Me.UCReportIO1.Add vbcombobox, 1, "Code", "StoreDescription", "0", "9", "Code", "Description", "SELECT * FROM tblpub_StoreKind Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "CalculatedGood", "0", "1", "Code", "Description", "SELECT * FROM tblpub_CalculatedGood Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "SortOrder", "0", "1", "Code", "Description", "SELECT * FROM tblpub_Sort Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
    
     Case "RepGarson"
        Me.UCReportIO1.MSG = "ê“«—‘ »œÂò«—Ì ê«—”Ê‰"
        ReportHeader = "ê“«—‘ »œÂò«—Ì ê«—”Ê‰"
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 2, "tFacm.[InCharge]", "garson", "1", "9999", "ppno", "nvcSurName", "SELECT * FROM tPer WHERE job = " & EnumIncharge.Garson & " And ActDeAct=1 Order By ppno"
        Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        
       
     Case "RepInventoryRecipt"
         
        Me.UCReportIO1.MSG = "   »— «”«” ‰Ê⁄ ›«ò Ê—"
       '  ReportHeader = "—”Ìœ «‰»«—"
       '  Me.UCReportIO1.AddForm = frmRep
        Me.UCReportIO1.Add vbcombobox, 1, "nvcDescription", "Status", "1", "99", "intStatusNo", "nvcDescription", "SELECT * FROM tStatusType ORDER BY intStatusNo", RightToLeft
        Me.UCReportIO1.Add vbtextbox, 1, "DiscountTotal", "FactorNo", "1", "999999", "No", "No", " Select Max No From tFacm Where AccountYear = dbo.Get_AccountYear()", RightToLeft
        Me.UCReportIO1.Add vbcombobox, 1, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
        
        For i = 0 To Me.UCReportIO1.Cmb(8).ListCount - 1
            Me.UCReportIO1.Cmb(8).ListIndex = i
            If AccountYear = Me.UCReportIO1.Cmb(8).Text Then
                Exit For
            End If
        Next

    Case "RepUsedGoodAmount"
        If clsStation.Language = Farsi Then
            Me.UCReportIO1.MSG = "ê“«—‘ „’—› „Ê«œ «Ê·ÌÂ"
            ReportHeader = "ê“«—‘ „’—› „Ê«œ «Ê·ÌÂ"
        Else
            Me.UCReportIO1.MSG = "First Material Consume Report"
            ReportHeader = "First Material Consume Report"
        End If
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 2, "InventoryNo", "Description", "1", "99", "InventoryNo", "Description", "SELECT * FROM tInventory Order By InventoryNo "
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        Me.UCReportIO1.Add vbcombobox, 2, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
    
    Case "RepLossGoodAmount"
        If clsStation.Language = Farsi Then
            Me.UCReportIO1.MSG = "ê“«—‘ ÷«Ì⁄« "
            ReportHeader = "ê“«—‘ ÷«Ì⁄« "
        Else
            Me.UCReportIO1.MSG = "Wastage Report "
            ReportHeader = "Wastage Report"
        End If
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 2, "GoodCode", "Good", "000000000", "999999999", "Code", "Name", "SELECT * FROM tGood Where GoodType <> 2 Order By Code "
        Me.UCReportIO1.Add vbcombobox, 2, "InventoryNo", "Description", "1", "99", "InventoryNo", "Description", "SELECT * FROM tInventory Order By InventoryNo "
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
    
     
        
    Case "RepCustomerList"
        
        Me.UCReportIO1.MSG = "·Ì”  „‘ —òÌ‰"
        ReportHeader = "·Ì”  „‘ —òÌ‰"
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbtextbox, 2, "Customer", "Customer", "1", "99999"
        Me.UCReportIO1.Add vbtextbox, 2, "sumprice", "SPrice", "0", "99999999999"
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        
    Case "RepStationSaleSummaryByUser", "RepStationSaleSummery"
    
        Me.UCReportIO1.MSG = "ê“«—‘ Œ·«’Â ›—Ê‘ ’‰œÊﬁœ«—«‰ - ﬂ«—»—«‰"
        ReportHeader = "ê“«—‘ Œ·«’Â ›—Ê‘ ’‰œÊﬁœ«—«‰ - ﬂ«—»—«‰"
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 2, "tFacm.[User]", "User", "1", "9999", "Uid", "nvcSurName", "SELECT * FROM tPer INNER JOIN tUser ON tPer.ppno = tUser.PPNo where ActDeAct=1 Order By tPer.ppno"
        Me.UCReportIO1.Add vbcombobox, 2, "tFacm.StationId", "StationId", "1", "999", "StationID", "StationID", "SELECT * FROM tStations  Order By StationId "
        Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        Me.UCReportIO1.Add vbcombobox, 2, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
    
    Case "RepGetOrderGoodAmount"
        Me.UCReportIO1.MSG = "ê“«—‘ «“ ﬂ«·«Â«Ì »Â ‰ﬁÿÂ ”›«—‘ —”ÌœÂ"
        ReportHeader = "ê“«—‘ «“ ﬂ«·«Â«Ì »Â ‰ﬁÿÂ ”›«—‘ —”ÌœÂ"
        Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
        Me.UCReportIO1.Add vbcombobox, 2, "level2", "L2", "1010", "9999", "Code", "Description", "SELECT * FROM tGoodLevel2 Order By Code "
        Me.UCReportIO1.Add vbcombobox, 2, "InventoryNo", "Description", "1", "99", "InventoryNo", "Description", "SELECT * FROM tInventory Order By InventoryNo "
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
     Case "RepCustomerLoan"
        If clsStation.Language = Farsi Then
            Me.UCReportIO1.MSG = "·Ì”  ›—Ê‘ «ﬁ”«ÿÌ ﬂ«—ﬂ‰«‰"
            ReportHeader = "·Ì”  ›—Ê‘ «ﬁ”«ÿÌ ﬂ«—ﬂ‰«‰"
        Else
            Me.UCReportIO1.MSG = "Payment By Instalment Report"
            ReportHeader = "Payment By Instalment Report"
        End If
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 2, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
       
    Case "RepDailyWeeding"
        Me.UCReportIO1.MSG = "ê“«—‘ „—«”„"
        ReportHeader = "ê“«—‘ „—«”„"
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        'Me.UCReportIO1.Add vbcombobox, 1, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
    Case "RepDailyPrize"
        Me.UCReportIO1.MSG = "ê“«—‘ Ã«Ì“Â"
        ReportHeader = "ê“«—‘ Ã«Ì“Â"
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        'Me.UCReportIO1.Add vbcombobox, 1, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
    Case "RepSaleShiftDailyPrize"
        Me.UCReportIO1.MSG = "ê“«—‘ ›—Ê‘ ‘Ì›  »«Ã«Ì“Â"
        ReportHeader = "ê“«—‘ ›—Ê‘ ‘Ì›  »«Ã«Ì“Â"
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        'Me.UCReportIO1.Add vbcombobox, 1, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
    
    Case "RepTableSellDetail"
        Me.UCReportIO1.MSG = "ê“«—‘ ›—Ê‘ „Ì“ Â«"
        ReportHeader = "ê“«—‘ ›—Ê‘ „Ì“ Â«"
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 2, "No", "Table", "1", "999", "No", "Name", "Select * From ttable Order By No", RightToLeft
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "Details", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Goods Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
     
     
     Case "RepCredit"
        Me.UCReportIO1.MSG = "·Ì”  »‰ Â«Ì œ—Ì«›  ‘œÂ"
        ReportHeader = "·Ì”  »‰ Â«Ì œ—Ì«›  ‘œÂ"
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 2, "tFacm.StationId", "StationId", "1", "999", "StationId", "StationId", "SELECT * FROM tStations  Order By StationId "
     
     Case "RepCheque"
        Me.UCReportIO1.MSG = "·Ì”  çﬂ Â«Ì œ—Ì«›  ‘œÂ"
        ReportHeader = "·Ì”  çﬂ Â«Ì œ—Ì«›  ‘œÂ"
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 2, "tFacm.StationId", "StationId", "1", "999", "StationId", "StationId", "SELECT * FROM tStations  Order By StationId "
     
     Case "RepCustomerBillPayment"
        Me.UCReportIO1.MSG = "ê“«—‘ ’Ê— Õ”«» „‘ —Ì«‰"
        ReportHeader = "ê“«—‘ ’Ê— Õ”«» „‘ —Ì«‰"
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbtextbox, 2, "Customer", "Customer", "1", "99999"
     
      Case "RepSubInventory"
         Me.UCReportIO1.MSG = "·Ì”  ›—Ê‘  €—›Â Â« »— «”«” ›Ì‘"
         ReportHeader = "·Ì”  ›—Ê‘  €—›Â Â« »— «”«” ›Ì‘"
         Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
         Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
         Me.UCReportIO1.Add vbcombobox, 2, "level2", "L2", "1010", "9999", "Code", "Description", "SELECT * FROM tGoodLevel2 Order By Code "
         Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
         Me.UCReportIO1.Add vbcombobox, 2, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
     Case "RepStationSale_CrossTab"
          Me.UCReportIO1.MSG = "ê“«—‘ ›—Ê‘ «Ì” ê«Â« »—«”«” €—›Â Â«"
          ReportHeader = "ê“«—‘ ›—Ê‘ «Ì” ê«Â« »—«”«” €—›Â Â«"
          Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
         ' Me.UCReportIO1.Add vbcombobox, 2, "tFacm.[User]", "User", "1", "9999", "ppno", "nvcSurName", "SELECT * FROM tPer INNER JOIN tUser ON tPer.ppno = tUser.PPNo Order By tPer.ppno"
          Me.UCReportIO1.Add vbcombobox, 2, "tFacm.StationId", "StationId", "1", "999", "StationId", "StationId", "SELECT * FROM tStations  Order By StationId "
          Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
          
        '  Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
          Me.UCReportIO1.Add vbcombobox, 2, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
    Case "RepBranchSale_CrossTab"
          Me.UCReportIO1.MSG = "ê“«—‘ ›—Ê‘ ‘⁄»Â Â« »—«”«” €—›Â Â«"
          ReportHeader = "ê“«—‘ ›—Ê‘ ‘⁄»Â Â« »—«”«” €—›Â Â«"
          Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
         ' Me.UCReportIO1.Add vbcombobox, 2, "tFacm.[User]", "User", "1", "9999", "ppno", "nvcSurName", "SELECT * FROM tPer INNER JOIN tUser ON tPer.ppno = tUser.PPNo Order By tPer.ppno"
          Me.UCReportIO1.Add vbcombobox, 2, "tFacm.StationId", "StationId", "1", "999", "StationId", "StationId", "SELECT * FROM tStations  Order By StationId "
          Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
          
        '  Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
          Me.UCReportIO1.Add vbcombobox, 2, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
    
    Case "RepTurnRecipt"
       
        Me.UCReportIO1.MSG = "  ê“«—‘ ê—œ‘ —”ÌœÂ«"
     '  ReportHeader = "—”Ìœ «‰»«—"
     '  Me.UCReportIO1.AddForm = frmRep
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 1, "nvcDescription", "Status", "1", "99", "intStatusNo", "nvcDescription", "SELECT * FROM tStatusType ORDER BY intStatusNo", RightToLeft
        Me.UCReportIO1.Add vbcombobox, 1, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
        Me.UCReportIO1.Add vbcombobox, 2, "level2", "L2", "1010", "9999", "Code", "Description", "SELECT * FROM tGoodLevel2 Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "
        
        
    Case "RepSellKindInfo_Bymonth"
       
        Me.UCReportIO1.MSG = "  ê“«—‘ ›—Ê‘ „«Â«‰Â »Â  ›òÌò ò«·«"
        Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
        Me.UCReportIO1.Add vbcombobox, 2, "Month", "Month", 1, 12, "Code", "Description", "SELECT * FROM tblPub_Month Order By code"
        Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
        
        For i = 0 To Me.UCReportIO1.Cmb(0).ListCount - 1
           Me.UCReportIO1.Cmb(0).ListIndex = i
           If AccountYear = Me.UCReportIO1.Cmb(0).Text Then
               Exit For
           End If
        Next
   
    Case "RepSellKindInfo"
        
        Me.UCReportIO1.MSG = "  ê“«—‘ ›—Ê‘ —Ê“«‰Â »Â  ›òÌò ò«·« "
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
         
    Case "RepInventoryAtomic_beneton"
       
        Me.UCReportIO1.MSG = " ›Â—”  „Õ’Ê·«  "
        Me.UCReportIO1.Add vbcombobox, 2, "InventoryNo", "Description", "1", "99", "InventoryNo", "Description", "SELECT * FROM tInventory Order By InventoryNo "
        Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
         
        For i = 0 To Me.UCReportIO1.Cmb(2).ListCount - 1
            Me.UCReportIO1.Cmb(2).ListIndex = i
            If AccountYear = Me.UCReportIO1.Cmb(2).Text Then
                Exit For
            End If
        Next
    
    Case "RepSelldaily"
       
        Me.UCReportIO1.MSG = " —Ê‰œ —Ê“«‰Â ò·Ì Ê Œ«·’ ›—Ê‘  "
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
         
    Case "RepSellKindInfo_ByYear"
       
        Me.UCReportIO1.MSG = " ê“«—‘ ”«·«‰Â ›—Ê‘ »Â  ›òÌò òœÌ‰ê ò«·« "
        Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
        Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
         
        For i = 0 To Me.UCReportIO1.Cmb(0).ListCount - 1
            Me.UCReportIO1.Cmb(0).ListIndex = i
            If AccountYear = Me.UCReportIO1.Cmb(0).Text Then
                Exit For
            End If
        Next
         
    Case "RepSellmonth"
       
        Me.UCReportIO1.MSG = "  —Ê‰œ „«ÂÌ«‰Â ò·Ì Ê Œ«·’ ›—Ê‘  "
        Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
        Me.UCReportIO1.Add vbcombobox, 2, "Month", "Month", 1, 12, "Code", "Description", "SELECT * FROM tblPub_Month Order By code"
        
        For i = 0 To Me.UCReportIO1.Cmb(0).ListCount - 1
            Me.UCReportIO1.Cmb(0).ListIndex = i
            If AccountYear = Me.UCReportIO1.Cmb(0).Text Then
                Exit For
            End If
        Next
         
    Case "RepSellinventory_Bymonth"
       
        Me.UCReportIO1.MSG = " ¬„«— „«ÂÌ«‰Â ›—Ê‘ ò«·« »Â  ›òÌò «‰»«— "
        Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
        Me.UCReportIO1.Add vbcombobox, 2, "Month", "Month", 1, 12, "Code", "Description", "SELECT * FROM tblPub_Month Order By code"
        Me.UCReportIO1.Add vbcombobox, 2, "InventoryNo", "Description", "1", "99", "InventoryNo", "Description", "SELECT * FROM tInventory Order By InventoryNo "
         
        For i = 0 To Me.UCReportIO1.Cmb(0).ListCount - 1
            Me.UCReportIO1.Cmb(0).ListIndex = i
            If AccountYear = Me.UCReportIO1.Cmb(0).Text Then
                Exit For
            End If
        Next

    Case "RepBuyinventory_Bymonth"
       
        Me.UCReportIO1.MSG = "¬„«— „«ÂÌ«‰Â Œ—Ìœ ò«·« »Â  ›òÌò «‰»«—"
        Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
        Me.UCReportIO1.Add vbcombobox, 2, "Month", "Month", 1, 12, "Code", "Description", "SELECT * FROM tblPub_Month Order By code"
        Me.UCReportIO1.Add vbcombobox, 2, "InventoryNo", "Description", "1", "99", "InventoryNo", "Description", "SELECT * FROM tInventory Order By InventoryNo "
         
        For i = 0 To Me.UCReportIO1.Cmb(0).ListCount - 1
            Me.UCReportIO1.Cmb(0).ListIndex = i
            If AccountYear = Me.UCReportIO1.Cmb(0).Text Then
                Exit For
            End If
        Next

     Case "RepBuymonth"
       
        Me.UCReportIO1.MSG = "  —Ê‰œ „«ÂÌ«‰Â ò·Ì Ê Œ«·’ Œ—Ìœ  "
        Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
        Me.UCReportIO1.Add vbcombobox, 2, "Month", "Month", 1, 12, "Code", "Description", "SELECT * FROM tblPub_Month Order By code"
    
        For i = 0 To Me.UCReportIO1.Cmb(0).ListCount - 1
            Me.UCReportIO1.Cmb(0).ListIndex = i
            If AccountYear = Me.UCReportIO1.Cmb(0).Text Then
                Exit For
            End If
        Next
    
 Case "RepSellBuyKindInfo"
       
        Me.UCReportIO1.MSG = "  —Ê‰œ —Ê“«‰Â Ê—Êœ Ê Œ—ÊÃ ò«·«  "
        Me.UCReportIO1.Add vbtextbox, 1, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 1, "InventoryNo", "Description", "1", "99", "InventoryNo", "Description", "SELECT * FROM tInventory Order By InventoryNo "
              
 Case "RepInventoryGood_Mojodi"
       
        Me.UCReportIO1.MSG = " ò«—œò” ò«·« "
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
        Me.UCReportIO1.Add vbcombobox, 1, "InventoryNo", "Description", "1", "99", "InventoryNo", "Description", "SELECT * FROM tInventory Order By InventoryNo "
        Me.UCReportIO1.Add vbcombobox, 1, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft

        For i = 0 To Me.UCReportIO1.Cmb(2).ListCount - 1
            Me.UCReportIO1.Cmb(2).ListIndex = i
            If AccountYear = Me.UCReportIO1.Cmb(2).Text Then
                Exit For
            End If
        Next
Case "RepMojodiYear"
        Me.UCReportIO1.MSG = "ê“«—‘ ⁄„·ò—œ  ⁄œ«œÌ «‰»«— - œ—  ”«· „«·Ì "
        Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
        Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
        Me.UCReportIO1.Add vbcombobox, 2, "InventoryNo", "Description", "1", "99", "InventoryNo", "Description", "SELECT * FROM tInventory Order By InventoryNo "
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "CalculatedGood", "0", "1", "Code", "Description", "SELECT * FROM tblpub_CalculatedGood Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "SortOrder", "0", "1", "Code", "Description", "SELECT * FROM tblpub_Sort Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
            
        For i = 0 To Me.UCReportIO1.Cmb(0).ListCount - 1
            Me.UCReportIO1.Cmb(0).ListIndex = i
            If AccountYear = Me.UCReportIO1.Cmb(0).Text Then
                Exit For
            End If
        Next
 Case "RepMojodiRialYear"
        Me.UCReportIO1.MSG = "ê“«—‘ ⁄„·ò—œ —Ì«·Ì «‰»«— - œ—  ”«· „«·Ì "
        Me.UCReportIO1.Add vbcombobox, 1, "AccountYear", "AccountYear", AccountYear, AccountYear, "AccountYear", "AccountYear", "SELECT * FROM tAccountYears Order By AccountYear"
        Me.UCReportIO1.Add vbcombobox, 2, "level1", "L1", "10", "99", "Code", "Description", "SELECT * FROM tGoodLevel1 Order By Code "
        Me.UCReportIO1.Add vbcombobox, 2, "InventoryNo", "Description", "1", "99", "InventoryNo", "Description", "SELECT * FROM tInventory Order By InventoryNo "
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "CalculatedGood", "0", "1", "Code", "Description", "SELECT * FROM tblpub_CalculatedGood Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "SortOrder", "0", "1", "Code", "Description", "SELECT * FROM tblpub_Sort Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "nvcBranchName", "Branch", "1", "99", "Branch", "nvcBranchName", "SELECT * FROM tBranch ORDER BY Branch", RightToLeft
    
       For i = 0 To Me.UCReportIO1.Cmb(0).ListCount - 1
            Me.UCReportIO1.Cmb(0).ListIndex = i
            If AccountYear = Me.UCReportIO1.Cmb(0).Text Then
                Exit For
            End If
        Next

Case "RepSeller"
        Me.UCReportIO1.MSG = "ê“«—‘ ⁄„·ﬂ—œ " & clsStation.SellerCaption
        ReportHeader = "ê“«—‘ ⁄„·ﬂ—œ " & clsStation.SellerCaption
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 2, "tFacm.[InCharge]", "Seller", "1", "9999", "pPNO", "nvcSurName", "SELECT * FROM tPer WHERE job = " & EnumIncharge.Seller & "Order By ppno"
        Me.UCReportIO1.Add vbcombobox, 1, "SellerReportType", "SellerReportType", "0", "1", "Code", "Description", "SELECT * FROM tblpub_SellerReportType Order By Code "

Case "RepAdditionalServices"
       Me.UCReportIO1.MSG = "ê“«—‘ ”—ÊÌ” Â«Ì œ—Ì«› Ì «“ „‘ —Ì«‰"
        ReportHeader = "ê“«—‘ ”—ÊÌ” Â«Ì œ—Ì«› Ì «“ „‘ —Ì«‰ "
        Me.UCReportIO1.Add vbtextbox, 2, "[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 2, "tFacm.[User]", "User", "1", "9999", "Uid", "nvcSurName", "SELECT * FROM tPer INNER JOIN tUser ON tPer.ppno = tUser.PPNo where ActDeAct=1 Order By tPer.ppno"
        Me.UCReportIO1.Add vbcombobox, 2, "tFacm.StationId", "StationId", "1", "999", "StationID", "StationID", "SELECT * FROM tStations  Order By StationId "
        Me.UCReportIO1.Add vbmaskbox, 2, "lTRIM([Time])", "Time", "00:00", "23:59"
        Me.UCReportIO1.Add vbcombobox, 1, "Paper", "PaperType", "0", "1", "Code", "Description", "SELECT * FROM tblPub_Paper Order By Code "

Case "RepCustomerBillPayment_Remain"
        Me.UCReportIO1.MSG = " ê“«—‘ ’Ê— Õ”«» „‘ —Ì«‰ »« „«‰œÂ «“ ﬁ»·"
        ReportHeader = "ê“«—‘ ’Ê— Õ”«» „‘ —Ì«‰"
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbtextbox, 2, "Customer", "Customer", "1", "99999"

Case "RepOPrder_ByDetail"
        Me.UCReportIO1.MSG = " ê“«—‘ ·Ì”  ”›«—‘« "
        ReportHeader = " ê“«—‘ ·Ì”  ”›«—‘« "
        Me.UCReportIO1.Add vbtextbox, 2, "tFacM.[Date]", "Dat", "70/01/01", "99/01/01"
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "OrderKind", "0", "1", "Code", "Description", "SELECT * FROM tblPub_OrderKind Order By Code "
        Me.UCReportIO1.Add vbcombobox, 1, "Code", "ShowKind", "0", "1", "Code", "Description", "SELECT * FROM tblPub_OrderDetail Order By Code "
        
End Select

beforeexit:
    
    
    Me.UCReportIO1.AddForm = frmRep
    Me.UCReportIO1.txt(0).Text = mvarDate '  Mid(ClsDate.shamsi(Date), 3, 8)
    Me.UCReportIO1.txt(1).Text = Mid(clsDate.shamsi(Date), 3, 8)
    If clsArya.MiladiDate = 1 Then
        FarDate1.Visible = False
        FarDate2.Visible = False
    Else
        FarDate1.Text = "13" + mvarDate
        FarDate2.Text = clsDate.shamsi(Date) ' Mid(ClsDate.shamsi(Date), 3, 8)
        FarDate1.Top = Me.UCReportIO1.txt(0).Top + 230
        FarDate1.Height = Me.UCReportIO1.txt(0).Height + 50
        FarDate1.Left = Me.UCReportIO1.txt(0).Left
        FarDate2.Top = Me.UCReportIO1.txt(1).Top + 230
        FarDate2.Height = Me.UCReportIO1.txt(1).Height + 50
        FarDate2.Left = Me.UCReportIO1.txt(1).Left
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Set clsDate = Nothing
    modgl.RightButton True
    Me.UCReportIO1.RemoveIndex
    
'    If ClsFormAccess.frmReports = True Then
'        frmReports.Show
'    End If
    
End Sub

'Private Sub Form_Resize()
'    On Error Resume Next
'     If Me.ScaleHeight > 0 Then
'        Me.Height = iHeight
'        Me.Width = iWidth
'     End If
'End Sub

Public Sub ExitForm()
    Unload Me
End Sub
Sub Set_RepCashInvoice()
        
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        User1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        User1 = -1
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        User2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        User2 = -1
    End If
    
    If Trim(frmRep.UCReportIO1.msk(6)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(6))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(7)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(7))
    Else
        Time2 = "23:59"
    End If

    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Station1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Station1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Station2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Station2 = 999
    End If
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    
    ReDim Parameter(11) As Parameter
    'ArrayUbound = 10
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@User1", adInteger, 4, User1)
    Parameter(7) = GenerateInputParameter("@User2", adInteger, 4, User2)
    Parameter(8) = GenerateInputParameter("@Time1", adVarWChar, 50, Time1)
    Parameter(9) = GenerateInputParameter("@Time2", adVarWChar, 50, Time2)
    Parameter(10) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
    Parameter(11) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)

    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCashInvoice.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCashInvoice_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepCashInvoice_En.rpt"
        Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepCashInvoice_En_A4.rpt"
        End If
    End If
End Sub
Sub Set_RepGroupGoodsSale() '0
 
    If frmRep.UCReportIO1.txt(2) = "" Then
       Sup1 = 1
    Else
       Sup1 = Me.UCReportIO1.txt(2).Text
    End If
    If frmRep.UCReportIO1.txt(3) = "" Then
       Sup2 = 9999
    Else
       Sup2 = Me.UCReportIO1.txt(3).Text
    End If
    
    If Trim(frmRep.UCReportIO1.msk(10)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(10))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(11)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(11))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Station1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Station1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Station2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Station2 = 999
    End If

    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        GoodLevel1FromCode = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex)
    Else
        GoodLevel1FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(7).Text <> "" Then
        GoodLevel1ToCode = frmRep.UCReportIO1.Cmb(7).ItemData(frmRep.UCReportIO1.Cmb(7).ListIndex)
    Else
        GoodLevel1ToCode = 999999999
    End If
    
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        GoodLevel2FromCode = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex)
    Else
        GoodLevel2FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(9).Text <> "" Then
        GoodLevel2ToCode = frmRep.UCReportIO1.Cmb(9).ItemData(frmRep.UCReportIO1.Cmb(9).ListIndex)
    Else
        GoodLevel2ToCode = 999999999
    End If
    If frmRep.UCReportIO1.Cmb(12).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(12).ItemData(frmRep.UCReportIO1.Cmb(12).ListIndex) - 1
    Else
        PaperType = 0
    End If
    If frmRep.UCReportIO1.Cmb(14).ItemData(frmRep.UCReportIO1.Cmb(14).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(14).ItemData(frmRep.UCReportIO1.Cmb(14).ListIndex)
    Else
        b1 = 0
    End If
    If frmRep.UCReportIO1.Cmb(15).ItemData(frmRep.UCReportIO1.Cmb(15).ListIndex) <> 0 Then
        b2 = frmRep.UCReportIO1.Cmb(15).ItemData(frmRep.UCReportIO1.Cmb(15).ListIndex)
    Else
        b2 = 0
    End If
    
        ReDim Parameter(18) As Parameter

        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
        Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        Parameter(6) = GenerateInputParameter("@Sup1", adInteger, 4, Sup1)
        Parameter(7) = GenerateInputParameter("@Sup2", adInteger, 4, Sup2)
        Parameter(8) = GenerateInputParameter("@Time1", adVarWChar, 50, Time1)
        Parameter(9) = GenerateInputParameter("@Time2", adVarWChar, 50, Time2)
        Parameter(10) = GenerateInputParameter("@Status", adInteger, 4, 2)
        Parameter(11) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
        Parameter(12) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)
        Parameter(13) = GenerateInputParameter("@FromGCodeL1", adInteger, 4, GoodLevel1FromCode)
        Parameter(14) = GenerateInputParameter("@ToGCodeL1", adInteger, 4, GoodLevel1ToCode)
        Parameter(15) = GenerateInputParameter("@FromGCodeL2", adInteger, 4, GoodLevel2FromCode)
        Parameter(16) = GenerateInputParameter("@ToGCodeL2", adInteger, 4, GoodLevel2ToCode)
        Parameter(17) = GenerateInputParameter("@Branch1", adInteger, 4, b1)
        Parameter(18) = GenerateInputParameter("@Branch2", adInteger, 4, b2)
        If clsStation.Language = Farsi Then
            If PaperType = 0 Then
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepGroupGoodsSale.rpt"
            Else
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepGroupGoodsSale_A4.rpt"
            End If
        Else
            If PaperType = 0 Then
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepGroupGoodsSale_En.rpt"
            Else
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepGroupGoodsSale_En_A4.rpt"
            End If
        End If

End Sub
  

Sub Set_RepFichSale()     '5

    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        User1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        User1 = -1
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        User2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        User2 = -1
    End If
    
    If Trim(frmRep.UCReportIO1.msk(6)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(6))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(7)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(7))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Station1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Station1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Station2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Station2 = 999
    End If
    
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        PaperType = 0
    End If

    
    
        ReDim Parameter(12) As Parameter

        'ArrayUbound = 12
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, " ”«⁄  : " & Mid(Str(Time), 1, 5))
        Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        Parameter(6) = GenerateInputParameter("@Status", adInteger, 4, 2)
        Parameter(7) = GenerateInputParameter("@User1", adInteger, 4, User1)
        Parameter(8) = GenerateInputParameter("@User2", adInteger, 4, User2)
        Parameter(9) = GenerateInputParameter("@Time1", adVarWChar, 50, Time1)
        Parameter(10) = GenerateInputParameter("@Time2", adVarWChar, 50, Time2)
        Parameter(11) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
        Parameter(12) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)
        If clsStation.Language = Farsi Then
            If PaperType = 0 Then
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepFichSale.rpt"
            Else
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepFichSale_A4.rpt"
            End If
        Else
            If PaperType = 0 Then
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepFichSale_En.rpt"
            Else
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepFichSale_En_A4.rpt"
            End If
        End If
End Sub
Sub Set_RepFichBuy()   '12

    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        User1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        User1 = -1
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        User2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        User2 = -1
    End If
    
    If Trim(frmRep.UCReportIO1.msk(6)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(6))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(7)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(7))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Station1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Station1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Station2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Station2 = 999
    End If
    
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
        ReDim Parameter(12) As Parameter
        'ArrayUbound = 12
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, " ”«⁄  : " & Mid(Str(Time), 1, 5))
        Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        Parameter(6) = GenerateInputParameter("@Status", adInteger, 4, 1)
        Parameter(7) = GenerateInputParameter("@User1", adInteger, 4, User1)
        Parameter(8) = GenerateInputParameter("@User2", adInteger, 4, User2)
        Parameter(9) = GenerateInputParameter("@Time1", adVarWChar, 50, Time1)
        Parameter(10) = GenerateInputParameter("@Time2", adVarWChar, 50, Time2)
        Parameter(11) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
        Parameter(12) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)
        If clsStation.Language = Farsi Then
            If PaperType = 0 Then
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepFichBuy.rpt"
            Else
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepFichBuy_A4.rpt"
            End If
        Else
            If PaperType = 0 Then
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepFichBuy_En.rpt"
            Else
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepFichBuy_En_A4.rpt"
            End If
        End If
End Sub

Sub Set_RepTimePerSell()
 
    If Trim(frmRep.UCReportIO1.msk(2)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(2))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(3)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(3))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) - 1
    Else
        PaperType = 0
    End If
    If frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex)
    Else
        b1 = 0
    End If
    If frmRep.UCReportIO1.Cmb(7).ItemData(frmRep.UCReportIO1.Cmb(7).ListIndex) <> 0 Then
        b2 = frmRep.UCReportIO1.Cmb(7).ItemData(frmRep.UCReportIO1.Cmb(7).ListIndex)
    Else
        b2 = 0
    End If
    
    ReDim Parameter(9) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@TimeBefore", adVarWChar, 50, Time1)
    Parameter(7) = GenerateInputParameter("@TimeAfter", adVarWChar, 50, Time2)
    Parameter(8) = GenerateInputParameter("@Branch1", adInteger, 4, b1)
    Parameter(9) = GenerateInputParameter("@Branch2", adInteger, 4, b2)
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepPercentInvoicePerHour.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepPercentInvoicePerHour_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepPercentInvoicePerHour_En.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepPercentInvoicePerHour_En_A4.rpt"
        End If
    End If
End Sub

Sub Set_RepGroupGoodsBuy()
    Dim s As String
    
    If frmRep.UCReportIO1.txt(2) = "" Then
       Sup1 = 1
    Else
       Sup1 = Me.UCReportIO1.txt(2).Text
    End If
    If frmRep.UCReportIO1.txt(3) = "" Then
       Sup2 = 9999
    Else
       Sup2 = Me.UCReportIO1.txt(3).Text
    End If
    
    If Trim(frmRep.UCReportIO1.msk(10)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(6))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(11)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(7))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Station1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Station1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Station2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Station2 = 999
    End If
    
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        GoodLevel1FromCode = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex)
    Else
        GoodLevel1FromCode = 0
    End If

    If frmRep.UCReportIO1.Cmb(7).Text <> "" Then
        GoodLevel1ToCode = frmRep.UCReportIO1.Cmb(7).ItemData(frmRep.UCReportIO1.Cmb(7).ListIndex)
    Else
        GoodLevel1ToCode = 999999999
    End If

    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        GoodLevel2FromCode = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex)
    Else
        GoodLevel2FromCode = 0
    End If

    If frmRep.UCReportIO1.Cmb(9).Text <> "" Then
        GoodLevel2ToCode = frmRep.UCReportIO1.Cmb(9).ItemData(frmRep.UCReportIO1.Cmb(9).ListIndex)
    Else
        GoodLevel2ToCode = 999999999
    End If
    
    If frmRep.UCReportIO1.Cmb(12).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(12).ItemData(frmRep.UCReportIO1.Cmb(12).ListIndex) - 1
    Else
        PaperType = 0
    End If
    If frmRep.UCReportIO1.Cmb(14).ItemData(frmRep.UCReportIO1.Cmb(14).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(14).ItemData(frmRep.UCReportIO1.Cmb(14).ListIndex)
    Else
        b1 = 0
    End If
    If frmRep.UCReportIO1.Cmb(15).ItemData(frmRep.UCReportIO1.Cmb(15).ListIndex) <> 0 Then
        b2 = frmRep.UCReportIO1.Cmb(15).ItemData(frmRep.UCReportIO1.Cmb(15).ListIndex)
    Else
        b2 = 0
    End If
    
        ReDim Parameter(18) As Parameter

        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, " ”«⁄  : " & Mid(Str(Time), 1, 5))
        Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        Parameter(6) = GenerateInputParameter("@Status", adInteger, 4, 1)
        Parameter(7) = GenerateInputParameter("@Sup1", adInteger, 4, Sup1)
        Parameter(8) = GenerateInputParameter("@Sup2", adInteger, 4, Sup2)
        Parameter(9) = GenerateInputParameter("@Time1", adVarWChar, 50, Time1)
        Parameter(10) = GenerateInputParameter("@Time2", adVarWChar, 50, Time2)
        Parameter(11) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
        Parameter(12) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)
        Parameter(13) = GenerateInputParameter("@FromGCodeL1", adInteger, 4, GoodLevel1FromCode)
        Parameter(14) = GenerateInputParameter("@ToGCodeL1", adInteger, 4, GoodLevel1ToCode)
        Parameter(15) = GenerateInputParameter("@FromGCodeL2", adInteger, 4, GoodLevel2FromCode)
        Parameter(16) = GenerateInputParameter("@ToGCodeL2", adInteger, 4, GoodLevel2ToCode)
        Parameter(17) = GenerateInputParameter("@Branch1", adInteger, 4, b1)
        Parameter(18) = GenerateInputParameter("@Branch2", adInteger, 4, b2)
       If clsStation.Language = Farsi Then
            If PaperType = 0 Then
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepGroupGoodsBuy.rpt"
            Else
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepGroupGoodsBuy_A4.rpt"
            End If
        Else
            If PaperType = 0 Then
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepGroupGoodsBuy_En.rpt"
            Else
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepGroupGoodsBuy_En_A4.rpt"
            End If
        End If
 End Sub
Sub Set_RepDetailGoods()   ' Case 2
    Dim s As String
    
    If frmRep.UCReportIO1.txt(2) = "" Then
       Sup1 = 1
    Else
       Sup1 = Me.UCReportIO1.txt(2).Text
    End If
    If frmRep.UCReportIO1.txt(3) = "" Then
       Sup2 = 9999
    Else
       Sup2 = Me.UCReportIO1.txt(3).Text
    End If
    
    If Trim(frmRep.UCReportIO1.msk(6)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(6))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(7)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(7))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Station1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Station1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Station2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Station2 = 999
    End If
 
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        PaperType = 0
    End If
       If frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex)
    Else
        b1 = 0
    End If
    If frmRep.UCReportIO1.Cmb(11).ItemData(frmRep.UCReportIO1.Cmb(11).ListIndex) <> 0 Then
        b2 = frmRep.UCReportIO1.Cmb(11).ItemData(frmRep.UCReportIO1.Cmb(11).ListIndex)
    Else
        b2 = 0
    End If
 
        ReDim Parameter(14) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
        Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        Parameter(6) = GenerateInputParameter("@Sup1", adInteger, 4, Sup1)
        Parameter(7) = GenerateInputParameter("@Sup2", adInteger, 4, Sup2)
        Parameter(8) = GenerateInputParameter("@Time1", adVarWChar, 50, Time1)
        Parameter(9) = GenerateInputParameter("@Time2", adVarWChar, 50, Time2)
        Parameter(10) = GenerateInputParameter("@Status", adInteger, 4, Statusvar)
        Parameter(11) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
        Parameter(12) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)
        Parameter(13) = GenerateInputParameter("@Branch1", adInteger, 4, b1)
        Parameter(14) = GenerateInputParameter("@Branch2", adInteger, 4, b2)
        
        If clsStation.Language = Farsi Then
            If PaperType = 0 Then
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepDetailGoods.rpt"
            Else
               If strCategory = "07" And strDelegate = "00" And clsArya.CustomerId = 104 Then
                    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepDetailGoods_Tejarat_A4.rpt"
               Else
                    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepDetailGoods_A4.rpt"
               End If
            End If
        Else
            If PaperType = 0 Then
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepDetailGoods_En.rpt"
            Else
               If strCategory = "07" And strDelegate = "00" And clsArya.CustomerId = 104 Then
                    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepDetailGoods_Tejarat_En_A4.rpt"
               Else
                    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepDetailGoods_En_A4.rpt"
               End If
            End If
        End If
End Sub
Sub Set_RepServeKindSell()    '   Case 6
    Dim s As String
    
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        User1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        User1 = -1
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        User2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        User2 = -1
    End If
    
    If Trim(frmRep.UCReportIO1.msk(6)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(6))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(7)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(7))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Station1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Station1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Station2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Station2 = 999
    End If
 
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        PaperType = 0
    End If
 
    
    ReDim Parameter(12) As Parameter

    'ArrayUbound = 12
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, " ”«⁄  : " & Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@Status", adInteger, 4, 2)
    Parameter(7) = GenerateInputParameter("@User1", adInteger, 4, User1)
    Parameter(8) = GenerateInputParameter("@User2", adInteger, 4, User2)
    Parameter(9) = GenerateInputParameter("@Time1", adVarWChar, 50, Time1)
    Parameter(10) = GenerateInputParameter("@Time2", adVarWChar, 50, Time2)
    Parameter(11) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
    Parameter(12) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)
                           
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepServeKindSell.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepServeKindSell_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepServeKindSell_En.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepServeKindSell_En_A4.rpt"
        End If

    End If
End Sub

Sub Set_RepBedeh()

    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        User1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        User1 = -1
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        User2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        User2 = -1
    End If
    
    If Trim(frmRep.UCReportIO1.msk(4)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(4))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(5)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(5))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    ReDim Parameter(9) As Parameter
    'ArrayUbound = 8
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@User1", adInteger, 4, User1)
    Parameter(7) = GenerateInputParameter("@User2", adInteger, 4, User2)
    Parameter(8) = GenerateInputParameter("@Time1", adVarWChar, 50, Time1)
    Parameter(9) = GenerateInputParameter("@Time2", adVarWChar, 50, Time2)
    
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCarrierBillPayment.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepCarrierBillPayment_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepCarrierBillPayment_En.rpt"
        Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepCarrierBillPayment_En_A4.rpt"
        End If
    End If
End Sub
Public Sub Set_RepGarson()

    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        User1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        User1 = -1
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        User2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        User2 = -1
    End If
    
    If Trim(frmRep.UCReportIO1.msk(4)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(4))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(5)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(5))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    ReDim Parameter(9) As Parameter
    'ArrayUbound = 8
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@User1", adInteger, 4, User1)
    Parameter(7) = GenerateInputParameter("@User2", adInteger, 4, User2)
    Parameter(8) = GenerateInputParameter("@Time1", adVarWChar, 50, Time1)
    Parameter(9) = GenerateInputParameter("@Time2", adVarWChar, 50, Time2)
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepGarsonBillPayment.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepGarsonBillPayment_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepGarsonBillPayment_En.rpt"
        Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepGarsonBillPayment_En_A4.rpt"
        End If
    End If
End Sub

Sub Set_RepBedehDetail()

    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        User1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        User1 = -1
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        User2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        User2 = -1
    End If
    
    If Trim(frmRep.UCReportIO1.msk(4)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(4))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(5)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(5))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    ReDim Parameter(9) As Parameter
    'ArrayUbound = 8
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@FromTime", adVarWChar, 50, Time1)
    Parameter(7) = GenerateInputParameter("@ToTime", adVarWChar, 50, Time2)
    Parameter(8) = GenerateInputParameter("@FromUser", adInteger, 4, User1)
    Parameter(9) = GenerateInputParameter("@ToUser", adInteger, 4, User2)
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCarrierDebitDetail.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepCarrierDebitDetail_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepCarrierDebitDetail_En.rpt"
        Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepCarrierDebitDetail_En_A4.rpt"
        End If
    End If
End Sub

Sub Set_RepCustPricDiscount()
    Dim TmpStatus As Integer
    If frmRep.UCReportIO1.txt(2) = "" Then
       cust1 = 1
    Else
       cust1 = Me.UCReportIO1.txt(2).Text
    End If
    If frmRep.UCReportIO1.txt(3) = "" Then
       cust2 = 9999999
    Else
       cust2 = Me.UCReportIO1.txt(3).Text
    End If
    If frmRep.UCReportIO1.txt(4) = "" Then
       SPrice1 = 0
    Else
       SPrice1 = Val(Me.UCReportIO1.txt(4).Text)
    End If
    If frmRep.UCReportIO1.txt(5) = "" Then
       SPrice2 = 999999999
    Else
       SPrice2 = Val(Me.UCReportIO1.txt(5).Text)
    End If
    If frmRep.UCReportIO1.txt(6) = "" Then
       Discont1 = 0
    Else
       Discont1 = Val(Me.UCReportIO1.txt(6).Text)
    End If
    If frmRep.UCReportIO1.txt(7) = "" Then
       Discont2 = 999999999
    Else
       Discont2 = Val(Me.UCReportIO1.txt(7).Text)
    End If
    
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        CustomerType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        CustomerType = 0
    End If
   
    If frmRep.UCReportIO1.Cmb(10).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex) - 1
    Else
        PaperType = 0
    End If
    If clsStation.Language = Farsi Then
        If CustomerType = 0 Then   ' Customers
           If Statusvar = 0 Then
               TmpStatus = 2 'Sale
               ReportHeader = " ·Ì”  ›—Ê‘ »Â „‘ —ﬂÌ‰ "
           Else
               TmpStatus = 5 'SaleReturn
               ReportHeader = "·Ì”  »—ê‘  «“ ›—Ê‘ „‘ —ﬂÌ‰"
           End If
        Else                    'Suppliers
           If Statusvar = 0 Then
               TmpStatus = 1 'Buy
               ReportHeader = " ·Ì”  Œ—Ìœ «“  «„Ì‰ ﬂ‰‰œê«‰ "
           Else
               TmpStatus = 4 'BuyReturn
               ReportHeader = "·Ì”  »—ê‘  «“ Œ—Ìœ «„Ì‰ ﬂ‰‰œê«‰"
           End If
        End If
    Else
        If CustomerType = 0 Then   ' Customers
            If Statusvar = 0 Then
                TmpStatus = 2 'Sale
                ReportHeader = " Customer Sale Report "
            Else
                TmpStatus = 5 'SaleReturn
                ReportHeader = "Customer Sale Revocation Report"
            End If
        Else                    'Suppliers
            If Statusvar = 0 Then
                TmpStatus = 1 'Buy
                ReportHeader = " Suppliers Purchase Report "
            Else
                TmpStatus = 4 'BuyReturn
                ReportHeader = "Suppliers Purchase Revocation Report"
            End If
        End If
    End If
    
    ReDim Parameter(12) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@FromDiscount", adDouble, 8, Discont1)
    Parameter(7) = GenerateInputParameter("@ToDiscount", adDouble, 8, Discont2)
    Parameter(8) = GenerateInputParameter("@FromSumPrice", adDouble, 8, SPrice1)
    Parameter(9) = GenerateInputParameter("@ToSumPrice", adDouble, 8, SPrice2)
    Parameter(10) = GenerateInputParameter("@FromCustCode", adInteger, 4, cust1)
    Parameter(11) = GenerateInputParameter("@ToCustCode", adInteger, 4, cust2)
    Parameter(12) = GenerateInputParameter("@Status", adInteger, 4, TmpStatus)
   If clsStation.Language = Farsi Then
        If CustomerType = 0 Then
            If PaperType = 0 Then
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCustomerBuyDiscount.rpt"
            Else
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepCustomerBuyDiscount_A4.rpt"
            End If
        Else
            If PaperType = 0 Then
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSupplierBuyDiscount.rpt"
            Else
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepSupplierBuyDiscount_A4.rpt"
            End If
        End If
    Else
        If CustomerType = 0 Then
            If PaperType = 0 Then
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepCustomerBuyDiscount_En.rpt"
            Else
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepCustomerBuyDiscount_En_A4.rpt"
            End If
        Else
            If PaperType = 0 Then
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepSupplierBuyDiscount_En.rpt"
            Else
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepSupplierBuyDiscount_En_A4.rpt"
            End If
        End If
    End If

End Sub

Sub Set_RepCustPrice()
    Dim TmpStatus As Integer
    If frmRep.UCReportIO1.txt(2) = "" Then
       cust1 = 1
    Else
       cust1 = Me.UCReportIO1.txt(2).Text
    End If
    If frmRep.UCReportIO1.txt(3) = "" Then
       cust2 = 999999
    Else
       cust2 = Me.UCReportIO1.txt(3).Text
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        CustomerType = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) - 1
    Else
        CustomerType = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
       DetailType = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
    Else
        DetailType = 0
    End If
 
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        PaperType = 0
    End If
    If clsStation.Language = Farsi Then
        If CustomerType = 0 Then   ' Customers
           If Statusvar = 0 Then
               TmpStatus = 2 'Sale
               ReportHeader = " ·Ì”  ›—Ê‘ »Â „‘ —ﬂÌ‰ - Ã“∆Ì« "
           Else
               TmpStatus = 5 'SaleReturn
               ReportHeader = "·Ì”  »—ê‘  «“ ›—Ê‘ „‘ —ﬂÌ‰- Ã“∆Ì« "
           End If
        Else                    'Suppliers
           If Statusvar = 0 Then
               TmpStatus = 1 'Buy
               ReportHeader = " ·Ì”  Œ—Ìœ «“  «„Ì‰ ﬂ‰‰œê«‰ - Ã“∆Ì« "
           Else
               TmpStatus = 4 'BuyReturn
               ReportHeader = "·Ì”  »—ê‘  «“ Œ—Ìœ «„Ì‰ ﬂ‰‰œê«‰- Ã“∆Ì« "
           End If
        End If
    Else
        If CustomerType = 0 Then   ' Customers
            If Statusvar = 0 Then
                TmpStatus = 2 'Sale
                ReportHeader = " Customer Sale Detail Report"
            Else
                TmpStatus = 5 'SaleReturn
                ReportHeader = "Customer Sale Revocation Detail Report"
            End If
        Else                    'Suppliers
            If Statusvar = 0 Then
                TmpStatus = 1 'Buy
                ReportHeader = "Suppliers Purchase Detail Report"
            Else
                TmpStatus = 4 'BuyReturn
                ReportHeader = "Suppliers Purchase Revocation Detail Report"
            End If
        End If
    End If
    
    ReDim Parameter(8) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@FromCustCode", adInteger, 4, cust1)
    Parameter(7) = GenerateInputParameter("@ToCustCode", adInteger, 4, cust2)
    Parameter(8) = GenerateInputParameter("@Status", adInteger, 4, TmpStatus)
    
    If clsStation.Language = Farsi Then
        If CustomerType = 0 Then
           If DetailType = 0 Then
              If PaperType = 0 Then
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCustomerBuyDetails.rpt"
              Else
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepCustomerBuyDetails_A4.rpt"
              End If
           Else
              If PaperType = 0 Then
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCustomerBuyDetails_Goods.rpt"
              Else
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepCustomerBuyDetails_Goods_A4.rpt"
              End If
           End If
    
        Else
           If DetailType = 0 Then
              If PaperType = 0 Then
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSupplierBuyDetails.rpt"
              Else
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepSupplierBuyDetails_A4.rpt"
              End If
           Else
              If PaperType = 0 Then
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSupplierBuyDetails_Goods.rpt"
              Else
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepSupplierBuyDetails_Goods_A4.rpt"
              End If
           End If
        End If
    Else
        If CustomerType = 0 Then
           If DetailType = 0 Then
              If PaperType = 0 Then
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepCustomerBuyDetails_En.rpt"
              Else
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepCustomerBuyDetails_En_A4.rpt"
              End If
           Else
              If PaperType = 0 Then
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepCustomerBuyDetails_Goods_En.rpt"
              Else
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepCustomerBuyDetails_Goods_En_A4.rpt"
              End If
           End If
    
        Else
           If DetailType = 0 Then
              If PaperType = 0 Then
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepSupplierBuyDetails_En.rpt"
              Else
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepSupplierBuyDetails_En_A4.rpt"
              End If
           Else
              If PaperType = 0 Then
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepSupplierBuyDetails_Goods_En.rpt"
              Else
                 CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepSupplierBuyDetails_Goods_En_A4.rpt"
              End If
           End If
        End If
    End If
End Sub
'---------------------------
'Sub Set_RepCustInfo()
'    Dim varAnswer As String
'    Load frmMsg
'    frmMsg.FWlblMsg.Caption = " . ¬Ì« Å—Ì‰ — ›—Ê‘ê«ÂÌ «” "
'    frmMsg.fwBtn(0).ButtonType = flwButtonOk
'    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
'    frmMsg.fwBtn(1).ButtonType = flwButtonNo
'    frmMsg.fwBtn(1).Caption = "Œ—ÊÃ"
'    frmMsg.Show vbModal
'    varAnswer = modgl.mvarMsgIdx
'    '---------------------------
'    Dim S As String
'    '------------------------Anahid
'    S = "SELECT '" & Mid(ClsDate.shamsi(Date), 3, 8) & "' AS SystemDate, ' " & _
'        ClsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)) & "' AS SystemDay, '" & _
'        " ”«⁄  : " & Mid(str(Time), 1, 5) & "' AS SystemTime, " & _
'        "tcust.Code, Name + ' ' + Family + ' ' + WorkName AS Name, " & _
'        "Address, CarryFee, PaykFee,tel1,credit,tDistance.Description FROM " & _
'        "tCust left outer join tDistance on tcust.Distance=tDistance.code WHERE " & _
'        Me.UCReportIO1.Sqlstr
'    '------------------------Anahid
'    Clipboard.Clear
'    Clipboard.SetText S
'
'    Cnn.Execute "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ViewCustInfo]') and OBJECTPROPERTY(id, N'IsView') = 1)drop view [dbo].[ViewCustInfo]"
'    Cnn.Execute "create view ViewCustInfo as " & S
'
'
'    '----------------------------
'    If varAnswer = vbYes Then
'
'        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCustInfo.rpt"
'
'    Else
'
'        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepCustInfo_A4.rpt"
'
'    End If
'
'End Sub

Sub Set_RepSystemGroup()

    If Trim(frmRep.UCReportIO1.msk(2)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(2))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(3)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(3))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) - 1
    Else
        PaperType = 0
    End If
    If frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex)
    Else
        b1 = 0
    End If
    If frmRep.UCReportIO1.Cmb(7).ItemData(frmRep.UCReportIO1.Cmb(7).ListIndex) <> 0 Then
        b2 = frmRep.UCReportIO1.Cmb(7).ItemData(frmRep.UCReportIO1.Cmb(7).ListIndex)
    Else
        b2 = 0
    End If
    
    ReDim Parameter(9) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@FromTime", adVarWChar, 50, Time1)
    Parameter(7) = GenerateInputParameter("@ToTime", adVarWChar, 50, Time2)
    Parameter(8) = GenerateInputParameter("@Branch1", adInteger, 4, b1)
    Parameter(9) = GenerateInputParameter("@Branch2", adInteger, 4, b2)
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSystemGroups.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepSystemGroups_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepSystemGroups_En.rpt"
        Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepSystemGroups_En_A4.rpt"
        End If
    End If
    
End Sub
Sub Set_RepSerialFich()

    If frmRep.UCReportIO1.txt(2).Text <> "" Then

        FactorNo1 = CInt(frmRep.UCReportIO1.txt(2).Text)
    Else
        FactorNo1 = 0
    End If
    
    If frmRep.UCReportIO1.txt(3).Text <> "" Then
    
        FactorNo2 = Val(frmRep.UCReportIO1.txt(3).Text)
    Else
        FactorNo2 = 9999999
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) - 1
    Else
        PaperType = 0
    End If
    ReDim Parameter(7) As Parameter
    'ArrayUbound = 6
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@FromFacNo", adInteger, 4, FactorNo1)
    Parameter(7) = GenerateInputParameter("@ToFacNo", adInteger, 4, FactorNo2)
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSerialDailySale.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepSerialDailySale_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepSerialDailySale_En.rpt"
        Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepSerialDailySale_En_A4.rpt"
        End If
    End If

End Sub

Sub Set_RepShift()

    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Shift1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Shift1 = 1
    End If
    
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Shift2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Shift2 = 1
    End If

    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        User1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        User1 = -1
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        User2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        User2 = -1
    End If

    If Trim(frmRep.UCReportIO1.msk(6)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(6))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(7)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(7))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    ReDim Parameter(11) As Parameter
    'ArrayUbound = 11

    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, " ”«⁄  : " & Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@FromTime", adVarWChar, 50, Time1)
    Parameter(7) = GenerateInputParameter("@ToTime", adVarWChar, 50, Time2)
    Parameter(8) = GenerateInputParameter("@FromShift", adInteger, 4, Shift1)
    Parameter(9) = GenerateInputParameter("@ToShift", adInteger, 4, Shift2)
    Parameter(10) = GenerateInputParameter("@FromUser", adInteger, 4, User1)
    Parameter(11) = GenerateInputParameter("@ToUser", adInteger, 4, User2)
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\ReShiftSale.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\ReShiftSale_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\ReShiftSale_En.rpt"
        Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\ReShiftSale_En_A4.rpt"
        End If
    End If

End Sub

Sub Set_RepMojodi()

If strCategory = "24" Then

    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        b1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        b1 = 0
    End If
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        b2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        b2 = 0
    End If
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Inventory1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Inventory1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        GoodLevel1FromCode = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex)
    Else
        GoodLevel1FromCode = 0
    End If
     If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        InventoryType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        InventoryType = 0
    End If
    If frmRep.UCReportIO1.Cmb(10).Text <> "" Then
        GoodFlag = frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex) - 1
    Else
        GoodFlag = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(12).Text <> "" Then
        CalculatedRialType = frmRep.UCReportIO1.Cmb(12).ItemData(frmRep.UCReportIO1.Cmb(12).ListIndex)
    Else
        CalculatedRialType = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(14).Text <> "" Then
        SortOrder = frmRep.UCReportIO1.Cmb(14).ItemData(frmRep.UCReportIO1.Cmb(14).ListIndex)
    Else
        SortOrder = 0
    End If
    
  
    ReDim Parameter(11) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Branch1", adInteger, 4, b1)
    Parameter(2) = GenerateInputParameter("@Branch2", adInteger, 4, b2)
    Parameter(3) = GenerateInputParameter("@InventoryNo", adInteger, 4, Inventory1)
    Parameter(4) = GenerateInputParameter("@FromGLvl1Code", adInteger, 4, GoodLevel1FromCode)
    Parameter(5) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(6) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(7) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(8) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(9) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(10) = GenerateInputParameter("@Flag", adTinyInt, 2, GoodFlag)
    Parameter(11) = GenerateInputParameter("@Sort", adInteger, 4, SortOrder)
        
    ReportHeader = "ê“«—‘ „ÊÃÊœÌ  ⁄œ«œÌ —Ì«·Ì «‰»«—"
    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepInventoryAtomicRials_Average_A4.rpt"
    
Else
    
'----------------------------
    
''    If frmRep.UCReportIO1.txt(2) = "" Then
''       Sup1 = 1
''    Else
''       Sup1 = Me.UCReportIO1.txt(2).Text
''    End If
''    If frmRep.UCReportIO1.txt(3) = "" Then
''       Sup2 = 9999
''    Else
''       Sup2 = Me.UCReportIO1.txt(3).Text
''    End If
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        GoodLevel1FromCode = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        GoodLevel1FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        GoodLevel1ToCode = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        GoodLevel1ToCode = 999999999
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Inventory1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Inventory1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Inventory2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Inventory2 = 1
    End If
    
''''     If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
''''        InventoryType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
''''    Else
''''        InventoryType = 0
''''    End If
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        GoodFlag = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
    Else
        GoodFlag = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        PaperType = 0
    End If
    If frmRep.UCReportIO1.Cmb(10).Text <> "" Then
        SortOrder = frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex)
    Else
        SortOrder = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(12).ItemData(frmRep.UCReportIO1.Cmb(12).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(12).ItemData(frmRep.UCReportIO1.Cmb(12).ListIndex)
    Else
        b1 = 0
    End If
    
    ReDim Parameter(11) As Parameter
    
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
'''        Parameter(1) = GenerateInputParameter("@Sup1", adInteger, 4, Sup1)
'''        Parameter(2) = GenerateInputParameter("@Sup2", adInteger, 4, Sup2)
        Parameter(1) = GenerateInputParameter("@FromLevel1", adInteger, 4, GoodLevel1FromCode)
        Parameter(2) = GenerateInputParameter("@ToLevel1", adInteger, 4, GoodLevel1ToCode)
        Parameter(3) = GenerateInputParameter("@InventoryNo1", adInteger, 4, Inventory1)
        Parameter(4) = GenerateInputParameter("@InventoryNo2", adInteger, 4, Inventory2)
        Parameter(5) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(6) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(7) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
        Parameter(8) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
'''        Parameter(11) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        Parameter(9) = GenerateInputParameter("@Flag", adTinyInt, 2, GoodFlag)
        Parameter(10) = GenerateInputParameter("@Branch", adInteger, 4, b1)
        Parameter(11) = GenerateInputParameter("@Sort", adInteger, 4, SortOrder)
        If clsStation.Language = Farsi Then
            ReportHeader = "ê“«—‘ „ÊÃÊœÌ  ⁄œ«œÌ «‰»«—"
            If PaperType = 0 Then
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepInventoryAtomicRemain.rpt"
            Else
               CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepInventoryAtomicRemain_A4.rpt"
            End If
        Else
            ReportHeader = "Number Inventory Report"
            If PaperType = 0 Then
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepInventoryAtomicRemain_En.rpt"
            Else
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepInventoryAtomicRemain_En_A4.rpt"
            End If
        End If
End If
End Sub
Sub Set_RepMojodiRial()
    
''''    If frmRep.UCReportIO1.txt(2) = "" Then
''''       Sup1 = 1
''''    Else
''''       Sup1 = Me.UCReportIO1.txt(2).Text
''''    End If
''''    If frmRep.UCReportIO1.txt(3) = "" Then
''''       Sup2 = 9999
''''    Else
''''       Sup2 = Me.UCReportIO1.txt(3).Text
''''    End If
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        GoodLevel1FromCode = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        GoodLevel1FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        GoodLevel1ToCode = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        GoodLevel1ToCode = 999999999
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Inventory1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Inventory1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Inventory2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Inventory2 = 1
    End If
    
''''     If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
''''        InventoryType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
''''    Else
''''        InventoryType = 0
''''    End If
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        GoodFlag = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
    Else
        GoodFlag = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(10).Text <> "" Then
        SortOrder = frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex)
    Else
        SortOrder = 0
    End If
    If frmRep.UCReportIO1.Cmb(12).ItemData(frmRep.UCReportIO1.Cmb(12).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(12).ItemData(frmRep.UCReportIO1.Cmb(12).ListIndex)
    Else
        b1 = 0
    End If
    
    ReDim Parameter(11) As Parameter
    
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
'''        Parameter(1) = GenerateInputParameter("@Sup1", adInteger, 4, Sup1)
'''        Parameter(2) = GenerateInputParameter("@Sup2", adInteger, 4, Sup2)
        Parameter(1) = GenerateInputParameter("@FromLevel1", adInteger, 4, GoodLevel1FromCode)
        Parameter(2) = GenerateInputParameter("@ToLevel1", adInteger, 4, GoodLevel1ToCode)
        Parameter(3) = GenerateInputParameter("@InventoryNo1", adInteger, 4, Inventory1)
        Parameter(4) = GenerateInputParameter("@InventoryNo2", adInteger, 4, Inventory2)
        Parameter(5) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(6) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(7) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
        Parameter(8) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
'''        Parameter(11) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        Parameter(9) = GenerateInputParameter("@Flag", adTinyInt, 2, GoodFlag)
        Parameter(10) = GenerateInputParameter("@Branch", adInteger, 4, b1)
        Parameter(11) = GenerateInputParameter("@Sort", adInteger, 4, SortOrder)
        If clsStation.Language = Farsi Then
            ReportHeader = "ê“«—‘ „ÊÃÊœÌ —Ì«·Ì «‰»«—"
            If PaperType = 0 Then
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepInventoryAtomicRials.rpt"
            Else
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepInventoryAtomicRials_A4.rpt"
            End If
        Else
            ReportHeader = "Remain Inventory Report"
            If PaperType = 0 Then
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepInventoryAtomicRials_En.rpt"
            Else
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepInventoryAtomicRials_En_A4.rpt"
            End If
        End If

End Sub
Sub Set_RepGoodList()
    
    If frmRep.UCReportIO1.Cmb(0).Text <> "" Then
        GoodLevel1FromCode = frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex)
    Else
        GoodLevel1FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(1).Text <> "" Then
        GoodLevel1ToCode = frmRep.UCReportIO1.Cmb(1).ItemData(frmRep.UCReportIO1.Cmb(1).ListIndex)
    Else
        GoodLevel1ToCode = 999999999
    End If
    
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        GoodLevel2FromCode = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        GoodLevel2FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        GoodLevel2ToCode = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        GoodLevel2ToCode = 999999999
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    
    ReDim Parameter(7) As Parameter
    'ArrayUbound = 7
   
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, " ”«⁄  : " & Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@FromGLvl1Code", adInteger, 4, GoodLevel1FromCode)
    Parameter(5) = GenerateInputParameter("@ToGLvl1Code", adInteger, 4, GoodLevel1ToCode)
    Parameter(6) = GenerateInputParameter("@FromGLvl2Code", adInteger, 4, GoodLevel2FromCode)
    Parameter(7) = GenerateInputParameter("@ToGLvl2Code", adInteger, 4, GoodLevel2ToCode)
    
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepGoodList.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepGoodList_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepGoodList_En.rpt"
         Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepGoodList_En_A4.rpt"
        End If
    End If

End Sub
Sub Set_RepStationSaleSummery()

''''    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
''''        User1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
''''    Else
''''        User1 = -1
''''    End If
''''
''''    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
''''        User2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
''''    Else
''''        User2 = -1
''''    End If
''''
''''
''''    If Trim(frmRep.UCReportIO1.msk(6)) <> ":" Then
''''        Time1 = Trim(frmRep.UCReportIO1.msk(6))
''''    Else
''''        Time1 = "00:00"
''''    End If
''''
''''    If Trim(frmRep.UCReportIO1.msk(7)) <> ":" Then
''''        Time2 = Trim(frmRep.UCReportIO1.msk(7))
''''    Else
''''        Time2 = "23:59"
''''    End If
''''
''''    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
''''        Station1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
''''    Else
''''        Station1 = 0
''''    End If
''''
''''    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
''''        Station2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
''''    Else
''''        Station2 = 999
''''    End If
''''
''''    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
''''        PaperType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
''''    Else
''''        PaperType = 0
''''    End If
''''    If frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex) <> 0 Then
''''        b1 = frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex)
''''    Else
''''        b1 = 0
''''    End If
''''    If frmRep.UCReportIO1.Cmb(11).ItemData(frmRep.UCReportIO1.Cmb(11).ListIndex) <> 0 Then
''''        b2 = frmRep.UCReportIO1.Cmb(11).ItemData(frmRep.UCReportIO1.Cmb(11).ListIndex)
''''    Else
''''        b2 = 0
''''    End If
''''
''''    ReDim Parameter(13) As Parameter
''''    'ArrayUbound = 10
''''    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
''''    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(ClsDate.shamsi(Date), 3, 8))
''''    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, ClsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
''''    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(str(Time), 1, 5))
''''    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
''''    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
''''    Parameter(6) = GenerateInputParameter("@FromTime", adVarWChar, 50, Time1)
''''    Parameter(7) = GenerateInputParameter("@ToTime", adVarWChar, 50, Time2)
''''    Parameter(8) = GenerateInputParameter("@FromUser", adInteger, 4, User1)
''''    Parameter(9) = GenerateInputParameter("@ToUser", adInteger, 4, User2)
''''    Parameter(10) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
''''    Parameter(11) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)
''''    Parameter(12) = GenerateInputParameter("@Branch1", adInteger, 4, b1)
''''    Parameter(13) = GenerateInputParameter("@Branch2", adInteger, 4, b2)
''''
''''    If clsStation.Language = Farsi Then
''''        If PaperType = 0 Then
''''           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepStationSaleSummery.rpt"
''''        Else
''''           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepStationSaleSummery_A4.rpt"
''''        End If
''''    Else
''''        If PaperType = 0 Then
''''            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepStationSaleSummery_En.rpt"
''''        Else
''''            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepStationSaleSummery_En_A4.rpt"
''''        End If
''''    End If
End Sub
Sub Set_RepStationSaleSummaryByUser()

    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        User1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        User1 = -1
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        User2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        User2 = -1
    End If


    If Trim(frmRep.UCReportIO1.msk(6)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(6))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(7)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(7))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Station1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Station1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Station2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Station2 = 999
    End If
    
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        PaperType = 0
    End If
    If frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex)
    Else
        b1 = 0
    End If
    If frmRep.UCReportIO1.Cmb(11).ItemData(frmRep.UCReportIO1.Cmb(11).ListIndex) <> 0 Then
        b2 = frmRep.UCReportIO1.Cmb(11).ItemData(frmRep.UCReportIO1.Cmb(11).ListIndex)
    Else
        b2 = 0
    End If
    
    ReDim Parameter(13) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@FromTime", adVarWChar, 50, Time1)
    Parameter(7) = GenerateInputParameter("@ToTime", adVarWChar, 50, Time2)
    Parameter(8) = GenerateInputParameter("@FromUser", adInteger, 4, User1)
    Parameter(9) = GenerateInputParameter("@ToUser", adInteger, 4, User2)
    Parameter(10) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
    Parameter(11) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)
    Parameter(12) = GenerateInputParameter("@Branch1", adInteger, 4, b1)
    Parameter(13) = GenerateInputParameter("@Branch2", adInteger, 4, b2)
    
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepStationSaleSummaryByUser.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepStationSaleSummaryByUser_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepStationSaleSummaryByUser_En.rpt"
        Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepStationSaleSummaryByUser_En_A4.rpt"
        End If
    End If
End Sub

Sub Set_RepCustomerList()


    If frmRep.UCReportIO1.txt(2) = "" Then
       cust1 = 1
    Else
       cust1 = Me.UCReportIO1.txt(2).Text
    End If
    If frmRep.UCReportIO1.txt(3) = "" Then
       cust2 = 999999
    Else
       cust2 = Me.UCReportIO1.txt(3).Text
    End If
    
    If frmRep.UCReportIO1.txt(4) = "" Then
       SPrice1 = 0
    Else
       SPrice1 = Val(Me.UCReportIO1.txt(4).Text)
    End If
    
    If frmRep.UCReportIO1.txt(5) = "" Then
       SPrice2 = 999999999
    Else
       SPrice2 = Val(Me.UCReportIO1.txt(5).Text)
    End If
    
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    ReDim Parameter(9) As Parameter
    'ArrayUbound = 12
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    
    Parameter(6) = GenerateInputParameter("@FromMaxBuy", adDouble, 8, SPrice1)
    Parameter(7) = GenerateInputParameter("@ToMaxBuy", adDouble, 8, SPrice2)
    Parameter(8) = GenerateInputParameter("@FromCustCode", adInteger, 4, cust1)
    Parameter(9) = GenerateInputParameter("@ToCustCode", adInteger, 4, cust2)
    
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCustomerList.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepCustomerList_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepCustomerList_En.rpt"
        Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepCustomerList_En_A4.rpt"
        End If
    End If
    
End Sub
Sub Set_RepGoodDifferences()
    
    If frmRep.UCReportIO1.Cmb(0).Text <> "" Then
        GoodLevel1FromCode = frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex)
    Else
        GoodLevel1FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(1).Text <> "" Then
        GoodLevel1ToCode = frmRep.UCReportIO1.Cmb(1).ItemData(frmRep.UCReportIO1.Cmb(1).ListIndex)
    Else
        GoodLevel1ToCode = 999999999
    End If
    
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        GoodLevel2FromCode = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        GoodLevel2FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        GoodLevel2ToCode = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        GoodLevel2ToCode = 999999999
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) - 1
    Else
        PaperType = 0
    End If
    ReDim Parameter(7) As Parameter
    'ArrayUbound = 7
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@FromGCodelvl1", adInteger, 4, GoodLevel1FromCode)
    Parameter(5) = GenerateInputParameter("@ToGCodelvl1", adInteger, 4, GoodLevel1ToCode)
    Parameter(6) = GenerateInputParameter("@FromGCodelvl2", adInteger, 4, GoodLevel2FromCode)
    Parameter(7) = GenerateInputParameter("@ToGCodelvl2", adInteger, 4, GoodLevel2ToCode)
    
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepGoodDifferences.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepGoodDifferences_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepGoodDifferences_En.rpt"
        Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepGoodDifferences_En_A4.rpt"
        End If
    End If
End Sub
Sub Set_RepInventoryRecipt()

    Dim Status1 As Integer
    Dim TempAccountYear As Integer
    
    If frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex) <> 0 Then
        Status1 = frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex)
    Else
        Status1 = 1
    End If
    
    If frmRep.UCReportIO1.txt(2).Text <> "" Then
        FactorNo1 = Val(frmRep.UCReportIO1.txt(2).Text)
    Else
        FactorNo1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        b1 = 1
    End If
    
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        TempAccountYear = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex)
    Else
        TempAccountYear = AccountYear
    End If
    
    ReDim Parameter(6) As Parameter
    'ArrayUbound = 5
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@FactorNo", adInteger, 4, FactorNo1)
    Parameter(2) = GenerateInputParameter("@Status", adInteger, 4, Status1)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, b1)
    Parameter(4) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(5) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(6) = GenerateInputParameter("@AccountYear", adSmallInt, 2, TempAccountYear)
    
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepInventoryReceipt.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepInventoryReceipt_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepInventoryReceipt_En.rpt"
        Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepInventoryReceipt_En_A4.rpt"
        End If
    End If
    
End Sub
Sub Set_RepUsedGoodAmount()

    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        Inventory1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        Inventory1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        Inventory2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        Inventory2 = 1
    End If
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) - 1
    Else
        PaperType = 0
    End If
    If frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex)
    Else
        b1 = 1
    End If
    
    ReDim Parameter(9) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@Datebefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@Type", adInteger, 4, EnumGoodType.forBuy)
    Parameter(7) = GenerateInputParameter("@InventoryNo1", adInteger, 4, Inventory1)
    Parameter(8) = GenerateInputParameter("@InventoryNo2", adInteger, 4, Inventory2)
    Parameter(9) = GenerateInputParameter("@Branch", adInteger, 4, b1)
    
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepGetUsedGoodAmount.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepGetUsedGoodAmount_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepGetUsedGoodAmount_En.rpt"
        Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepGetUsedGoodAmount_En_A4.rpt"
        End If
    End If

End Sub
Sub Set_RepLossGoodAmount()
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        GoodType1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        GoodType1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        GoodType2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        GoodType2 = 999999999
    End If
    
''''    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
''''        StoreType1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
''''    Else
''''        InventoryType = 1
''''    End If
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Inventory1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Inventory1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Inventory2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Inventory2 = 1
    End If
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
    Else
        PaperType = 0
    End If
    If frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex)
    Else
        b1 = 1
    End If
    
   
    ReDim Parameter(10) As Parameter
    
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@GoodType1", adInteger, 4, GoodType1)
        Parameter(2) = GenerateInputParameter("@GoodType2", adInteger, 4, GoodType2)
        Parameter(3) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(4) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(5) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
        Parameter(6) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(7) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        Parameter(8) = GenerateInputParameter("@Inventory1", adInteger, 4, Inventory1)
        Parameter(9) = GenerateInputParameter("@Inventory2", adInteger, 4, Inventory2)
        Parameter(10) = GenerateInputParameter("@Branch", adInteger, 4, b1)
    
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepGetLossGoodAmount.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepGetLossGoodAmount_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepGetLossGoodAmount_En.rpt"
        Else
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepGetLossGoodAmount_En_A4.rpt"
        End If
    End If

End Sub
Sub Set_RepGetOrderGoodAmount()

    If frmRep.UCReportIO1.Cmb(0).Text <> "" Then
        GoodLevel1FromCode = frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex)
    Else
        GoodLevel1FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(1).Text <> "" Then
        GoodLevel1ToCode = frmRep.UCReportIO1.Cmb(1).ItemData(frmRep.UCReportIO1.Cmb(1).ListIndex)
    Else
        GoodLevel1ToCode = 999999999
    End If
    
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        GoodLevel2FromCode = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        GoodLevel2FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        GoodLevel2ToCode = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        GoodLevel2ToCode = 999999999
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Inventory1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Inventory1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Inventory2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Inventory2 = 1
    End If
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    ReDim Parameter(9) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@FromGClvl1", adInteger, 4, GoodLevel1FromCode)
    Parameter(5) = GenerateInputParameter("@ToGClvl1", adInteger, 4, GoodLevel1ToCode)
    Parameter(6) = GenerateInputParameter("@FromGClvl2", adInteger, 4, GoodLevel2FromCode)
    Parameter(7) = GenerateInputParameter("@ToGClvl2", adInteger, 4, GoodLevel2ToCode)
    Parameter(8) = GenerateInputParameter("@Inventory1", adInteger, 4, Inventory1)
    Parameter(9) = GenerateInputParameter("@Inventory2", adInteger, 4, Inventory2)
    
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepGetOrderGoodAmount.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepGetOrderGoodAmount_A4.rpt"
         End If
    Else
        If PaperType = 0 Then
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepGetOrderGoodAmount_En.rpt"
        Else
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepGetOrderGoodAmount_En_A4.rpt"
        End If
    End If

End Sub
Sub Set_RepCustomerLoan()

    If frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        b1 = 0
    End If
    If frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex) <> 0 Then
        b2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        b2 = 0
    End If
    
    
    ReDim Parameter(7) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@Branch1", adInteger, 4, b1)
    Parameter(7) = GenerateInputParameter("@Branch2", adInteger, 4, b2)
    If clsStation.Language = Farsi Then
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepCustomerLoan_A4.rpt"
    Else
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepCustomerLoan_En_A4.rpt"
    End If

End Sub
Sub Set_RepDailyWeeding()
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex) - 1
    Else
        PaperType = 0
    End If
''    If frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) <> 0 Then
''        b1 = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex)
''    Else
''        b1 = 1
''    End If
    
    ReDim Parameter(5) As Parameter
    
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
        Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        'Parameter(10) = GenerateInputParameter("@Branch", adInteger, 4, b1)
    If clsStation.Language = Farsi Then
       If PaperType = 0 Then
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepDailyWeeding.rpt"
       Else
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepDailyWeeding_A4.rpt"
       End If
    Else
        If PaperType = 0 Then
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepDailyWeeding_En.rpt"
       Else
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepDailyWeeding_En_A4.rpt"
       End If
    End If

End Sub
Sub Set_RepDailyPrize()
      If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex) - 1
    Else
        PaperType = 0
    End If
''    If frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) <> 0 Then
''        b1 = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex)
''    Else
''        b1 = 1
''    End If
    
    ReDim Parameter(5) As Parameter
    
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
        Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        'Parameter(10) = GenerateInputParameter("@Branch", adInteger, 4, b1)
    If clsStation.Language = Farsi Then
       If PaperType = 0 Then
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepDailyPrize.rpt"
       Else
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepDailyPrize_A4.rpt"
       End If
    Else
        If PaperType = 0 Then
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepDailyPrize_En.rpt"
       Else
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepDailyPrize_En_A4.rpt"
       End If
    End If

End Sub
Sub Set_RepSaleShiftDailyPrize()
      If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex) - 1
    Else
        PaperType = 0
    End If
''    If frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) <> 0 Then
''        b1 = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex)
''    Else
''        b1 = 1
''    End If
    
    ReDim Parameter(5) As Parameter
    
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
        Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        'Parameter(10) = GenerateInputParameter("@Branch", adInteger, 4, b1)
    If clsStation.Language = Farsi Then
       If PaperType = 0 Then
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSaleShiftDailyPrize.rpt"
       Else
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepSaleShiftDailyPrize_A4.rpt"
       End If
    Else
        If PaperType = 0 Then
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepSaleShiftDailyPrize_En.rpt"
       Else
          CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepSaleShiftDailyPrize_En_A4.rpt"
       End If
    End If

End Sub
Sub Set_ReptableSellDetail()
    
    Dim TableNo1, TableNo2 As Integer
    
    If frmRep.UCReportIO1.Cmb(2) = "" Then
       TableNo1 = 1
    Else
       TableNo1 = UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    End If
    If frmRep.UCReportIO1.Cmb(3) = "" Then
       TableNo2 = 999
    Else
       TableNo2 = UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    End If
    
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
       DetailType = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) - 1
    Else
        DetailType = 0
    End If
 
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    ReDim Parameter(7) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@FromTableNo", adInteger, 4, TableNo1)
    Parameter(7) = GenerateInputParameter("@ToTableNo", adInteger, 4, TableNo2)
    
    If clsStation.Language = Farsi Then
       If DetailType = 0 Then
          If PaperType = 0 Then
             CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepTableSellDetail.rpt"
          Else
             CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTableSellDetail_A4.rpt"
          End If
       Else
          If PaperType = 0 Then
             CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepTableSellDetail_Goods.rpt"
          Else
             CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTableSellDetail_Goods_A4.rpt"
          End If
       End If

    Else
        If DetailType = 0 Then
           If PaperType = 0 Then
              CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepTableSellDetail_En.rpt"
           Else
              CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepTableSellDetail_En_A4.rpt"
           End If
        Else
           If PaperType = 0 Then
              CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepTableSellDetail_Goods_En.rpt"
           Else
              CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepTableSellDetail_Goods_En_A4.rpt"
           End If
        End If
    
    End If


End Sub

Sub Set_RepCheque() '0
 
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        Station1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        Station1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        Station2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        Station2 = 999
    End If

    
        ReDim Parameter(6) As Parameter

        Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
        Parameter(3) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(4) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        Parameter(5) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
        Parameter(6) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)

       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepChequeSellDetail_A4.rpt"
        
End Sub
Sub Set_RepCredit() '0
 
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        Station1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        Station1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        Station2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        Station2 = 999
    End If

    
        ReDim Parameter(6) As Parameter

        Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
        Parameter(3) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(4) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        Parameter(5) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
        Parameter(6) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)

       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCreditSellDetail.rpt"
        
End Sub
Sub Set_RepCustomerBillPayment()


    If frmRep.UCReportIO1.txt(2) = "" Then
       cust1 = 1
    Else
       cust1 = Me.UCReportIO1.txt(2).Text
    End If
    If frmRep.UCReportIO1.txt(3) = "" Then
       cust2 = 999999
    Else
       cust2 = Me.UCReportIO1.txt(3).Text
    End If
    
    
    ReDim Parameter(6) As Parameter
    'ArrayUbound = 12
    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(3) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(4) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    
    Parameter(5) = GenerateInputParameter("@FromCustCode", adInteger, 4, cust1)
    Parameter(6) = GenerateInputParameter("@ToCustCode", adInteger, 4, cust2)
    
    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCustomerBillPayment.rpt"
    
End Sub
Sub Set_RepSubInventory() '0
 
    
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        GoodLevel1FromCode = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        GoodLevel1FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        GoodLevel1ToCode = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        GoodLevel1ToCode = 999999999
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        GoodLevel2FromCode = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        GoodLevel2FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        GoodLevel2ToCode = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        GoodLevel2ToCode = 999999999
    End If
    If Trim(frmRep.UCReportIO1.msk(6)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(6))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(7)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(7))
    Else
        Time2 = "23:59"
    End If
    

    If frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex)
    Else
        b1 = 0
    End If
    If frmRep.UCReportIO1.Cmb(9).ItemData(frmRep.UCReportIO1.Cmb(9).ListIndex) <> 0 Then
        b2 = frmRep.UCReportIO1.Cmb(9).ItemData(frmRep.UCReportIO1.Cmb(9).ListIndex)
    Else
        b2 = 0
    End If
    
        ReDim Parameter(12) As Parameter

        Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
        Parameter(3) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(4) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        Parameter(5) = GenerateInputParameter("@Time1", adVarWChar, 50, Time1)
        Parameter(6) = GenerateInputParameter("@Time2", adVarWChar, 50, Time2)
        Parameter(7) = GenerateInputParameter("@FromGCodeL1", adInteger, 4, GoodLevel1FromCode)
        Parameter(8) = GenerateInputParameter("@ToGCodeL1", adInteger, 4, GoodLevel1ToCode)
        Parameter(9) = GenerateInputParameter("@FromGCodeL2", adInteger, 4, GoodLevel2FromCode)
        Parameter(10) = GenerateInputParameter("@ToGCodeL2", adInteger, 4, GoodLevel2ToCode)
        Parameter(11) = GenerateInputParameter("@Branch1", adInteger, 4, b1)
        Parameter(12) = GenerateInputParameter("@Branch2", adInteger, 4, b2)

        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSubInventoryByFich.rpt"

End Sub
Sub Set_RepStationSale_CrossTab() '0
 
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        Station1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        Station1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        Station2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        Station2 = 999
    End If

    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        GoodLevel1FromCode = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        GoodLevel1FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        GoodLevel1ToCode = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        GoodLevel1ToCode = 999999999
    End If
    If frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex)
    Else
        b1 = 0
    End If
    If frmRep.UCReportIO1.Cmb(7).ItemData(frmRep.UCReportIO1.Cmb(7).ListIndex) <> 0 Then
        b2 = frmRep.UCReportIO1.Cmb(7).ItemData(frmRep.UCReportIO1.Cmb(7).ListIndex)
    Else
        b2 = 0
    End If
    
        ReDim Parameter(7) As Parameter

        Parameter(0) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(1) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        Parameter(2) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
        Parameter(3) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)
        Parameter(4) = GenerateInputParameter("@FromGCodeL1", adInteger, 4, GoodLevel1FromCode)
        Parameter(5) = GenerateInputParameter("@ToGCodeL1", adInteger, 4, GoodLevel1ToCode)
        Parameter(6) = GenerateInputParameter("@Branch1", adInteger, 4, b1)
        Parameter(7) = GenerateInputParameter("@Branch2", adInteger, 4, b2)

        
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepStationSale_CrossTab_A4.rpt"
        
End Sub
Sub Set_RepBranchSale_CrossTab() '0
 
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        Station1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        Station1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        Station2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        Station2 = 999
    End If

    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        GoodLevel1FromCode = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        GoodLevel1FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        GoodLevel1ToCode = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        GoodLevel1ToCode = 999999999
    End If
    If frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex)
    Else
        b1 = 0
    End If
    If frmRep.UCReportIO1.Cmb(7).ItemData(frmRep.UCReportIO1.Cmb(7).ListIndex) <> 0 Then
        b2 = frmRep.UCReportIO1.Cmb(7).ItemData(frmRep.UCReportIO1.Cmb(7).ListIndex)
    Else
        b2 = 0
    End If
    
        ReDim Parameter(7) As Parameter

        Parameter(0) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
        Parameter(1) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
        Parameter(2) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
        Parameter(3) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)
        Parameter(4) = GenerateInputParameter("@FromGCodeL1", adInteger, 4, GoodLevel1FromCode)
        Parameter(5) = GenerateInputParameter("@ToGCodeL1", adInteger, 4, GoodLevel1ToCode)
        Parameter(6) = GenerateInputParameter("@Branch1", adInteger, 4, b1)
        Parameter(7) = GenerateInputParameter("@Branch2", adInteger, 4, b2)

        
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepBranchSale_CrossTab_A4.rpt"
        
End Sub
Sub Set_RepTurnRecipt()

    Dim Status1 As Integer
    If frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex) <> 0 Then
        Status1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        Status1 = 1
    End If
    
    If frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        b1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        GoodLevel2FromCode = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex)
    Else
        GoodLevel2FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(7).Text <> "" Then
        GoodLevel2ToCode = frmRep.UCReportIO1.Cmb(7).ItemData(frmRep.UCReportIO1.Cmb(7).ListIndex)
    Else
        GoodLevel2ToCode = 999999999
    End If
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    ReDim Parameter(8) As Parameter
    'ArrayUbound = 5
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(3) = GenerateInputParameter("@Status", adInteger, 4, Status1)
    Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, b1)
    Parameter(5) = GenerateInputParameter("@FromGCodeL2", adInteger, 4, GoodLevel2FromCode)
    Parameter(6) = GenerateInputParameter("@ToGCodeL2", adInteger, 4, GoodLevel2ToCode)
    Parameter(7) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(8) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    
    If PaperType = 0 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepTurnRecipt.rpt"
    Else
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTurnRecipt_A4.rpt"
    End If
    
End Sub
Sub Set_RepSellKindInfo_Bymonth()

    
    If frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex) <> 0 Then
        AccountYear = frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex)
   Else
  ''     Status1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex) <> 0 Then
        MonthBefore = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
   Else
       MonthBefore = 1
    End If
     If frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex) <> 0 Then
        MonthAfter = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
     Else
       MonthAfter = 12
    End If
    
  
    
    ReDim Parameter(5) As Parameter
    'ArrayUbound = 5
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@MonthBefore", adInteger, 4, MonthBefore)
    Parameter(2) = GenerateInputParameter("@MonthAfter", adInteger, 4, MonthAfter)
    Parameter(3) = GenerateInputParameter("@Year", adInteger, 4, AccountYear)
    Parameter(4) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(5) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    
   '' If PaperType = 0 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSellKindInfo_Bymonth.rpt"
   '' Else
    ''   CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTurnRecipt_A4.rpt"
   '' End If
    
End Sub
Sub Set_RepSellKindInfo()
  
    
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        FromGCodeL1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        FromGCodeL1 = 11
    End If
    
  
     If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        toGCodeL1 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        toGCodeL1 = 99
    End If
    
    
    ReDim Parameter(6) As Parameter
    'ArrayUbound = 5
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(3) = GenerateInputParameter("@FromGCodeL1", adInteger, 4, FromGCodeL1)
    Parameter(4) = GenerateInputParameter("@ToGCodeL1", adInteger, 4, toGCodeL1)
    Parameter(5) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(6) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    
    
   '' If PaperType = 0 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSellKindInfo.rpt"
   '' Else
    ''   CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTurnRecipt_A4.rpt"
   '' End If
    
End Sub
Sub Set_RepInventoryAtomic_beneton()

    
    If frmRep.UCReportIO1.Cmb(0).Text <> "" Then
        Inventory1 = frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex)
    Else
        Inventory1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(1).Text <> "" Then
        Inventory2 = frmRep.UCReportIO1.Cmb(1).ItemData(frmRep.UCReportIO1.Cmb(1).ListIndex)
    Else
        Inventory2 = 99
    End If
   If frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex) <> 0 Then
        AccountYear1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
   Else
        AccountYear1 = AccountYear
    End If
    
    
    
    
    ReDim Parameter(4) As Parameter
    'ArrayUbound = 5
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@InventoryNo1", adInteger, 4, Inventory1)
    Parameter(1) = GenerateInputParameter("@InventoryNo2", adInteger, 4, Inventory2)
    Parameter(2) = GenerateInputParameter("@AccountYear", adInteger, 4, AccountYear1)
    Parameter(3) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(4) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    
   '' If PaperType = 0 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepInventoryAtomic_beneton.rpt"
   '' Else
    ''   CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTurnRecipt_A4.rpt"
   '' End If
    
End Sub
Sub Set_RepSelldaily()

    
    
    ReDim Parameter(4) As Parameter
    'ArrayUbound = 5
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(3) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(4) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    
   '' If PaperType = 0 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSelldaily.rpt"
   '' Else
    ''   CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTurnRecipt_A4.rpt"
   '' End If
    
End Sub
Sub Set_RepSellKindInfo_ByYear()

    
    If frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex) <> 0 Then
        AccountYear = frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex)
   Else
  ''     Status1 = 1
    End If
  
    
    
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        FromGCodeL1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        FromGCodeL1 = 11
    End If
    
  
     If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        toGCodeL1 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        toGCodeL1 = 99
    End If
    
    
    
    ReDim Parameter(5) As Parameter
    'ArrayUbound = 5
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Year", adInteger, 4, AccountYear)
    Parameter(2) = GenerateInputParameter("@FromGCodeL1", adInteger, 4, FromGCodeL1)
    Parameter(3) = GenerateInputParameter("@ToGCodeL1", adInteger, 4, toGCodeL1)
    Parameter(4) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(5) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    
   '' If PaperType = 0 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSellKindInfo_ByYear.rpt"
   '' Else
    ''   CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTurnRecipt_A4.rpt"
   '' End If
    
End Sub

Sub Set_RepSellmonth()

    
    If frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex) <> 0 Then
        AccountYear = frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex)
   Else
  ''     Status1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex) <> 0 Then
        MonthBefore = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
   Else
       MonthBefore = 1
    End If
     If frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex) <> 0 Then
        MonthAfter = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
     Else
       MonthAfter = 12
    End If
    
    
 
    
    
    
    ReDim Parameter(5) As Parameter
    'ArrayUbound = 5
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@MonthBefore", adInteger, 4, MonthBefore)
    Parameter(2) = GenerateInputParameter("@MonthAfter", adInteger, 4, MonthAfter)
    Parameter(3) = GenerateInputParameter("@Year", adInteger, 4, AccountYear)
    Parameter(4) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(5) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    
   '' If PaperType = 0 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSellmonth.rpt"
   '' Else
    ''   CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTurnRecipt_A4.rpt"
   '' End If
    
End Sub

Sub Set_RepSellinventory_Bymonth()

    
    If frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex) <> 0 Then
        AccountYear = frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex)
   Else
  ''     Status1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex) <> 0 Then
        MonthBefore = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
   Else
       MonthBefore = 1
    End If
     If frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex) <> 0 Then
        MonthAfter = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
     Else
       MonthAfter = 12
    End If
    
     
    If frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) <> 0 Then
        Inventory1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Inventory1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex) <> 0 Then
        Inventory2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Inventory2 = 99
    End If
  
    
    
    ReDim Parameter(7) As Parameter
    'ArrayUbound = 5
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@MonthBefore", adInteger, 4, MonthBefore)
    Parameter(2) = GenerateInputParameter("@MonthAfter", adInteger, 4, MonthAfter)
    Parameter(3) = GenerateInputParameter("@Year", adInteger, 4, AccountYear)
    Parameter(4) = GenerateInputParameter("@FromInventory", adInteger, 4, Inventory1)
    Parameter(5) = GenerateInputParameter("@ToInventory", adInteger, 4, Inventory2)
    Parameter(6) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(7) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    
   '' If PaperType = 0 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSellinventory_Bymonth.rpt"
   '' Else
    ''   CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTurnRecipt_A4.rpt"
   '' End If
    
End Sub

Sub Set_RepBuyinventory_Bymonth()

    
     If frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex) <> 0 Then
        AccountYear = frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex)
   Else
  ''     Status1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex) <> 0 Then
        MonthBefore = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
   Else
       MonthBefore = 1
    End If
     If frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex) <> 0 Then
        MonthAfter = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
     Else
       MonthAfter = 12
    End If
    
     
    If frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) <> 0 Then
        Inventory1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Inventory1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex) <> 0 Then
        Inventory2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Inventory2 = 99
    End If
    
    
    
   ReDim Parameter(7) As Parameter
    'ArrayUbound = 5
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@MonthBefore", adInteger, 4, MonthBefore)
    Parameter(2) = GenerateInputParameter("@MonthAfter", adInteger, 4, MonthAfter)
    Parameter(3) = GenerateInputParameter("@Year", adInteger, 4, AccountYear)
    Parameter(4) = GenerateInputParameter("@FromInventory", adInteger, 4, Inventory1)
    Parameter(5) = GenerateInputParameter("@ToInventory", adInteger, 4, Inventory2)
    Parameter(6) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(7) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    
   '' If PaperType = 0 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepBuyinventory_Bymonth.rpt"
   '' Else
    ''   CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTurnRecipt_A4.rpt"
   '' End If
    
End Sub
Sub Set_RepBuymonth()

    
      If frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex) <> 0 Then
        AccountYear = frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex)
   Else
  ''     Status1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex) <> 0 Then
        MonthBefore = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
   Else
       MonthBefore = 1
    End If
     If frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex) <> 0 Then
        MonthAfter = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
     Else
       MonthAfter = 12
    End If
    
    
    
   ReDim Parameter(5) As Parameter
    'ArrayUbound = 5
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@MonthBefore", adInteger, 4, MonthBefore)
    Parameter(2) = GenerateInputParameter("@MonthAfter", adInteger, 4, MonthAfter)
    Parameter(3) = GenerateInputParameter("@Year", adInteger, 4, AccountYear)
    Parameter(4) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(5) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    
   '' If PaperType = 0 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepBuymonth.rpt"
   '' Else
    ''   CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTurnRecipt_A4.rpt"
   '' End If
    
End Sub
Sub Set_RepSellBuyKindInfo()

       
    If frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex) <> 0 Then
        Inventory1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        Inventory1 = 1
    End If

    
    
    ReDim Parameter(1) As Parameter
    'ArrayUbound = 5
     Parameter(0) = GenerateInputParameter("@Dateafter", adVarWChar, 50, DateBefor)
     Parameter(1) = GenerateInputParameter("@Inventory", adInteger, 4, Inventory1)
   '' If PaperType = 0 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSellBuyKindInfo.rpt"
   '' Else
    ''   CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTurnRecipt_A4.rpt"
   '' End If
    
End Sub
Sub Set_RepInventoryGood_Mojodi()

    
    If frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex) <> 0 Then
        AccountYear = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
   Else
  ''     Status1 = 1
    End If
           
    If frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) <> 0 Then
        Inventory1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Inventory1 = 1
    End If
     If frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex)
    Else
        b1 = 1
    End If
    ReDim Parameter(4) As Parameter
    'ArrayUbound = 5
    
    
    Parameter(0) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(1) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(2) = GenerateInputParameter("@Intinventoryno", adInteger, 4, Inventory1)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, b1)
    Parameter(4) = GenerateInputParameter("@AccountYear", adInteger, 4, AccountYear)
   '' If PaperType = 0 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepInventoryGood_Mojodi.rpt"
   '' Else
    ''   CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepTurnRecipt_A4.rpt"
   '' End If
    
End Sub
Sub Set_RepMojodiYear()
    

   If frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex) <> 0 Then
        AccountYear = frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex)
   Else
  ''     Status1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        GoodLevel1FromCode = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        GoodLevel1FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        GoodLevel1ToCode = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        GoodLevel1ToCode = 999999999
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Inventory1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Inventory1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Inventory2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Inventory2 = 1
    End If
    
''''     If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
''''        InventoryType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
''''    Else
''''        InventoryType = 0
''''    End If
''''    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
''''        GoodFlag = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
''''    Else
''''        GoodFlag = 0
''''    End If
    
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        SortOrder = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex)
    Else
        SortOrder = 0
    End If
    If frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex)
    Else
        b1 = 0
    End If
    
    ReDim Parameter(12) As Parameter
    
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@FromLevel1", adInteger, 4, GoodLevel1FromCode)
        Parameter(2) = GenerateInputParameter("@ToLevel1", adInteger, 4, GoodLevel1ToCode)
        Parameter(3) = GenerateInputParameter("@FromInventoryNo", adInteger, 4, Inventory1)
        Parameter(4) = GenerateInputParameter("@ToInventoryNo", adInteger, 4, Inventory2)
        Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, b1)
        Parameter(6) = GenerateInputParameter("@Accountyear", adInteger, 4, AccountYear)
        Parameter(7) = GenerateInputParameter("@CheckNotZeroMojodi ", adInteger, 4, 0)
        Parameter(8) = GenerateInputParameter("@CheckOrder ", adInteger, 4, 0)
        Parameter(9) = GenerateInputParameter("@SortItem", adInteger, 4, SortOrder)
        Parameter(10) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(11) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(12) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
   
'''        Parameter(11) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    
        
        
        If clsStation.Language = Farsi Then
            ReportHeader = "ê“«—‘ „ÊÃÊœÌ —Ì«·Ì «‰»«—"
''            If PaperType = 0 Then
''                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepInventoryAtomicRials.rpt"
''            Else
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepInventoryAtomicRemainYear_A4.rpt"
''            End If
        Else
            ReportHeader = "Remain Inventory Report"
            If PaperType = 0 Then
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepInventoryAtomicRials_En.rpt"
            Else
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepInventoryAtomicRials_En_A4.rpt"
            End If
        End If

End Sub
Sub Set_RepMojodiRialYear()
    
''''    If frmRep.UCReportIO1.txt(2) = "" Then
''''       Sup1 = 1
''''    Else
''''       Sup1 = Me.UCReportIO1.txt(2).Text
''''    End If
''''    If frmRep.UCReportIO1.txt(3) = "" Then
''''       Sup2 = 9999
''''    Else
''''       Sup2 = Me.UCReportIO1.txt(3).Text
''''    End If
   If frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex) <> 0 Then
        AccountYear = frmRep.UCReportIO1.Cmb(0).ItemData(frmRep.UCReportIO1.Cmb(0).ListIndex)
   Else
  ''     Status1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        GoodLevel1FromCode = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        GoodLevel1FromCode = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        GoodLevel1ToCode = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        GoodLevel1ToCode = 999999999
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Inventory1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Inventory1 = 1
    End If
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Inventory2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Inventory2 = 1
    End If
    
''''     If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
''''        InventoryType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
''''    Else
''''        InventoryType = 0
''''    End If
''''    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
''''        GoodFlag = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
''''    Else
''''        GoodFlag = 0
''''    End If
    
    If frmRep.UCReportIO1.Cmb(6).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(6).ItemData(frmRep.UCReportIO1.Cmb(6).ListIndex) - 1
    Else
        PaperType = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        SortOrder = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex)
    Else
        SortOrder = 0
    End If
    If frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex) <> 0 Then
        b1 = frmRep.UCReportIO1.Cmb(10).ItemData(frmRep.UCReportIO1.Cmb(10).ListIndex)
    Else
        b1 = 0
    End If
    
    ReDim Parameter(12) As Parameter
    
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@FromLevel1", adInteger, 4, GoodLevel1FromCode)
        Parameter(2) = GenerateInputParameter("@ToLevel1", adInteger, 4, GoodLevel1ToCode)
        Parameter(3) = GenerateInputParameter("@FromInventoryNo", adInteger, 4, Inventory1)
        Parameter(4) = GenerateInputParameter("@ToInventoryNo", adInteger, 4, Inventory2)
        Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, b1)
        Parameter(6) = GenerateInputParameter("@Accountyear", adInteger, 4, AccountYear)
        Parameter(7) = GenerateInputParameter("@CheckNotZeroMojodi ", adInteger, 4, 0)
        Parameter(8) = GenerateInputParameter("@CheckOrder ", adInteger, 4, 0)
        Parameter(9) = GenerateInputParameter("@SortItem", adInteger, 4, SortOrder)
        Parameter(10) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(11) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(12) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
   
'''        Parameter(11) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    
        
        
        If clsStation.Language = Farsi Then
            ReportHeader = "ê“«—‘ „ÊÃÊœÌ —Ì«·Ì «‰»«—"
''            If PaperType = 0 Then
''                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepInventoryAtomicRials.rpt"
''            Else
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepInventoryAtomicRialsYear_A4.rpt"
''            End If
        Else
            ReportHeader = "Remain Inventory Report"
            If PaperType = 0 Then
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepInventoryAtomicRials_En.rpt"
            Else
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepInventoryAtomicRials_En_A4.rpt"
            End If
        End If

End Sub
Sub Set_Repseller()


    If frmRep.UCReportIO1.Cmb(2).Text = "" Then
       cust1 = 1
    Else
       cust1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    End If
    If frmRep.UCReportIO1.Cmb(3).Text = "" Then
       cust2 = 999999
    Else
       cust2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        RepSellerType = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex) - 1
    Else
        RepSellerType = 0
    End If
    
    ReDim Parameter(6) As Parameter
    'ArrayUbound = 12
    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(3) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(4) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(5) = GenerateInputParameter("@FromSeller", adInteger, 4, cust1)
    Parameter(6) = GenerateInputParameter("@ToSeller", adInteger, 4, cust2)
    
    If RepSellerType = 0 Then
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepsellerList.rpt"
    Else
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSellerByFactor.rpt"
    End If
End Sub
Sub Set_RepAdditionalServices()    '   Case 6
    Dim s As String
    
    If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        User1 = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        User1 = -1
    End If
    
    If frmRep.UCReportIO1.Cmb(3).Text <> "" Then
        User2 = frmRep.UCReportIO1.Cmb(3).ItemData(frmRep.UCReportIO1.Cmb(3).ListIndex)
    Else
        User2 = -1
    End If
    
    If Trim(frmRep.UCReportIO1.msk(6)) <> ":" Then
        Time1 = Trim(frmRep.UCReportIO1.msk(6))
    Else
        Time1 = "00:00"
    End If
    
    If Trim(frmRep.UCReportIO1.msk(7)) <> ":" Then
        Time2 = Trim(frmRep.UCReportIO1.msk(7))
    Else
        Time2 = "23:59"
    End If
    
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        Station1 = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        Station1 = 0
    End If
    
    If frmRep.UCReportIO1.Cmb(5).Text <> "" Then
        Station2 = frmRep.UCReportIO1.Cmb(5).ItemData(frmRep.UCReportIO1.Cmb(5).ListIndex)
    Else
        Station2 = 999
    End If
 
    If frmRep.UCReportIO1.Cmb(8).Text <> "" Then
        PaperType = frmRep.UCReportIO1.Cmb(8).ItemData(frmRep.UCReportIO1.Cmb(8).ListIndex) - 1
    Else
        PaperType = 0
    End If
 
    
    ReDim Parameter(11) As Parameter

    'ArrayUbound = 12
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, " ”«⁄  : " & Mid(Str(Time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    Parameter(6) = GenerateInputParameter("@User1", adInteger, 4, User1)
    Parameter(7) = GenerateInputParameter("@User2", adInteger, 4, User2)
    Parameter(8) = GenerateInputParameter("@Time1", adVarWChar, 50, Time1)
    Parameter(9) = GenerateInputParameter("@Time2", adVarWChar, 50, Time2)
    Parameter(10) = GenerateInputParameter("@FromStationID", adInteger, 4, Station1)
    Parameter(11) = GenerateInputParameter("@ToStationID", adInteger, 4, Station2)
                           
    If clsStation.Language = Farsi Then
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepAdditionalServices.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepAdditionalServices_A4.rpt"
        End If
    Else
        If PaperType = 0 Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepAdditionalServices_En.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\A4\RepAdditionalServices_A4.rpt"
        End If

    End If
End Sub
Sub Set_RepCustomerBillPayment_Remain()


    If frmRep.UCReportIO1.txt(2) = "" Then
       cust1 = 1
    Else
       cust1 = Me.UCReportIO1.txt(2).Text
    End If
    If frmRep.UCReportIO1.txt(3) = "" Then
       cust2 = 999999
    Else
       cust2 = Me.UCReportIO1.txt(3).Text
    End If
    
    
    ReDim Parameter(6) As Parameter
    'ArrayUbound = 12
    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(3) = GenerateInputParameter("@DateBefore", adVarWChar, 50, DateBefor)
    Parameter(4) = GenerateInputParameter("@DateAfter", adVarWChar, 50, DateAfter)
    
    Parameter(5) = GenerateInputParameter("@FromCustCode", adInteger, 4, cust1)
    Parameter(6) = GenerateInputParameter("@ToCustCode", adInteger, 4, cust2)
    
    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCustomerBillPayment_Remain.rpt"
    
End Sub
Sub Set_RepOPrder_ByDetail()


   If frmRep.UCReportIO1.Cmb(2).Text <> "" Then
        OrderKind = frmRep.UCReportIO1.Cmb(2).ItemData(frmRep.UCReportIO1.Cmb(2).ListIndex)
    Else
        OrderKind = 0
    End If
    If frmRep.UCReportIO1.Cmb(4).Text <> "" Then
        ShowKind = frmRep.UCReportIO1.Cmb(4).ItemData(frmRep.UCReportIO1.Cmb(4).ListIndex)
    Else
        ShowKind = 0
    End If

    
    
    ReDim Parameter(6) As Parameter
    'ArrayUbound = 12
    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
    Parameter(3) = GenerateInputParameter("@FromDate", adVarWChar, 50, DateBefor)
    Parameter(4) = GenerateInputParameter("@ToDate", adVarWChar, 50, DateAfter)
    Parameter(5) = GenerateInputParameter("@Balance", adInteger, 4, OrderKind)
    Parameter(6) = GenerateInputParameter("@ShowKind", adInteger, 4, ShowKind)
    
    
    
    If ShowKind = 1 Then
      CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepOrderByDetail_A4.rpt"
    Else
      CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A4\RepOrderByDetail_Good_A4.rpt"
    End If
    
End Sub


Public Sub UCReportIO1_CommANDclick(Index As Integer)
    On Error GoTo ErrorHandler
    Dim strCnn As String
    Dim s, S1, S2 As String
    
    If frmRep.UCReportIO1.txt(0) <> "" Then
       DateBefor = frmRep.UCReportIO1.txt(0)
    Else
        If clsArya.MiladiDate = 0 Then
            DateBefor = "70/01/01"
        Else
            DateBefor = "01/01/01"
        End If
    End If
    
    If frmRep.UCReportIO1.txt(1) <> "" Then
       DateAfter = frmRep.UCReportIO1.txt(1)
    Else
        If clsArya.MiladiDate = 0 Then
            DateAfter = "99/12/30"
        Else
            DateAfter = "99/12/31"
        End If
    End If
    Select Case CRepFlag
        Case "RepGroupGoodsSale"
            Set_RepGroupGoodsSale
        Case "RepTimePerSell"
            Set_RepTimePerSell
        Case "RepFichSale"
            Set_RepFichSale
        Case "RepFichBuy"
            Set_RepFichBuy
        Case "RepDetailGoodsSale"
            Set_RepDetailGoods
        Case "RepServeKindSell"
            Set_RepServeKindSell
        Case "RepBedeh"
            Set_RepBedeh
        Case "RepBedehDetail"
            Set_RepBedehDetail
        Case "RepCustPricDiscount"
            Set_RepCustPricDiscount
        Case "RepCustPrice"
            Set_RepCustPrice
'        Case "RepCustInfo"
'            Set_RepCustInfo
        Case "RepSystemGroup"
            Set_RepSystemGroup
        Case "RepSerialFich"
            Set_RepSerialFich
        Case "RepGroupGoodsBuy"
            Set_RepGroupGoodsBuy
        Case "RepDetailGoodsBuy"
            Set_RepDetailGoods
        Case "RepShift"
            Set_RepShift
        Case "RepMojodi"
            Set_RepMojodi
        Case "RepMojodiRial"
            Set_RepMojodiRial
        Case "RepGarson"
            Set_RepGarson
        Case "RepCashInvoice"
            Set_RepCashInvoice
        Case "RepGoodList"
            Set_RepGoodList
        Case "RepStationSaleSummery"
            Set_RepStationSaleSummaryByUser  '  Set_RepStationSaleSummery
        Case "RepCustomerList"
            Set_RepCustomerList
        Case "RepInventoryRecipt"
            Set_RepInventoryRecipt
        Case "RepGoodDifferences"
            Set_RepGoodDifferences
        Case "RepUsedGoodAmount"
            Set_RepUsedGoodAmount
        Case "RepLossGoodAmount"
            Set_RepLossGoodAmount
        Case "RepStationSaleSummaryByUser"
            Set_RepStationSaleSummaryByUser
        Case "RepGetOrderGoodAmount"
            Set_RepGetOrderGoodAmount
        Case "RepDetailGoodsSaleReturn"
            Set_RepDetailGoods
        Case "RepDetailGoodsBuyReturn"
            Set_RepDetailGoods
        Case "RepCustPricDiscountReturn"
            Set_RepCustPricDiscount
        Case "RepCustPriceReturn"
            Set_RepCustPrice
        Case "RepCustomerLoan"
            Set_RepCustomerLoan
        Case "RepDailyWeeding"
            Set_RepDailyWeeding
        Case "RepDailyPrize"
            Set_RepDailyPrize
        Case "RepSaleShiftDailyPrize"
            Set_RepSaleShiftDailyPrize
        Case "RepTableSellDetail"
            Set_ReptableSellDetail
        Case "RepCredit"
            Set_RepCredit
        Case "RepCheque"
            Set_RepCheque
        Case "RepCustomerBillPayment"
            Set_RepCustomerBillPayment
        Case "RepSubInventory"
            Set_RepSubInventory
        Case "RepStationSale_CrossTab"
            Set_RepStationSale_CrossTab
        Case "RepBranchSale_CrossTab"
            Set_RepBranchSale_CrossTab
        Case "RepTurnRecipt"
           Set_RepTurnRecipt
        Case "RepSellKindInfo_Bymonth"
           Set_RepSellKindInfo_Bymonth
        Case "RepSellKindInfo"
           Set_RepSellKindInfo
        Case "RepInventoryAtomic_beneton"
           Set_RepInventoryAtomic_beneton
        Case "RepSelldaily"
           Set_RepSelldaily
        Case "RepSellKindInfo_ByYear"
           Set_RepSellKindInfo_ByYear
        Case "RepSellmonth"
           Set_RepSellmonth
        Case "RepSellinventory_Bymonth"
            Set_RepSellinventory_Bymonth
        Case "RepBuyinventory_Bymonth"
            Set_RepBuyinventory_Bymonth
        Case "RepBuymonth"
            Set_RepBuymonth
        Case "RepSellBuyKindInfo"
            Set_RepSellBuyKindInfo
        Case "RepInventoryGood_Mojodi"
            Set_RepInventoryGood_Mojodi
        Case "RepMojodiYear"
           Set_RepMojodiYear
        Case "RepMojodiRialYear"
           Set_RepMojodiRialYear
        Case "RepSeller"
           Set_Repseller
        Case "RepAdditionalServices"
            Set_RepAdditionalServices
        Case "RepCustomerBillPayment_Remain"
            Set_RepCustomerBillPayment_Remain
        Case "RepOPrder_ByDetail"
            Set_RepOPrder_ByDetail
    End Select
    '-----------------------
   ' CrystalReport1.ReportTitle = clsArya.StationName
    CrystalReport1.ReportTitle = ReportHeader
    CrystalReport1.Destination = crptToWindow 'crptToPrinter '
    Dim intIndex As Integer
    
    'Õ–› Å«—«„ —Â«Ì »Â Ã« „«‰œÂ «“ ê“«—‘ ﬁ»·Ì
    For intIndex = 0 To 100
        CrystalReport1.ParameterFields(intIndex) = ""
    Next intIndex
    
    '«›“Êœ‰ Å«—«„ —Â«Ì ê“«—‘ «‰ Œ«» ‘œÂ
    For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
        CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
    Next intIndex
   
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Connect = CrystallConnection
    CrystalReport1.Action = 1
    If PaperType = 1 Then
       CrystalReport1.PageZoom (100)
    Else
       CrystalReport1.PageZoom (100)
       
    End If
    Exit Sub
ErrorHandler:
   MsgBox err.Description & "  File Name:  " & CrystalReport1.ReportFileName
       Resume Next
   
End Sub

