VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClnReportBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClsReportBase"
Attribute VB_Ext_KEY = "Member0" ,"ClsReportBase"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mvarRc As New ADODB.Recordset
Private ClsGl As New ClsGl
Private ObjNewMember As New ClsReportBase

Public Function Add(DscrFull As String, DscrIn As String, DscrOut As String, Optional Frmt As String, Optional Key As String, Optional DataType As String) As ClsReportBase
    ObjNewMember.DscrFull = DscrFull
    ObjNewMember.DscrIn = DscrIn
    ObjNewMember.DscrOut = DscrOut
    ObjNewMember.Frmt = Frmt
    ObjNewMember.Key = Key
    ObjNewMember.DataType = DataType
    If Len(Key) = 0 Then
        mcol.Add ObjNewMember
    Else
        mcol.Add ObjNewMember, Key
    End If
    Set Add = ObjNewMember
    Set ObjNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClsReportBase
Attribute Item.VB_UserMemId = 0
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property


Public Sub remove(vntIndexKey As Variant)

    mcol.remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcol.[_NewEnum]
End Property


Private Sub Class_Initialize()
'creates the collection when this class is created
Set mcol = New Collection
Me.Add " «—ÌŒ", "«“  «—ÌŒ", " «  «—ÌŒ", , "Dat", "String"
Me.Add "›«ﬂ Ê—", "«“ ›«ﬂ Ê—", " « ›«ﬂ Ê—", , "Fac", "Integer"
Me.Add "’‰œÊﬁ", "«“ ’‰œÊﬁ", " « ’‰œÊﬁ", , "StationId", "Integer"
Me.Add "„‘ —ﬂ", "«“ «‘ —«ﬂ", " « «‘ —«ﬂ", , "Customer", "Integer"
Me.Add " «„Ì‰ ﬂ‰‰œÂ", "«“  «„Ì‰ ﬂ‰‰œÂ", " «  «„Ì‰ ﬂ‰‰œÂ", , "Supplier", "Integer"
Me.Add "ﬂ«—»—", "«“ ﬂ«—»—", " « ﬂ«—»—", , "User", "String"
Me.Add "”«⁄ ", "«“ ”«⁄ ", " « ”«⁄ ", , "Time", "String"
Me.Add " Œ›Ì›", "«“ Œ›Ì›", " «  Œ›Ì›", , "SDiscount", "Double"
Me.Add "Œ—Ìœ", "«“Œ—Ìœ", " «Œ—Ìœ", , "SPrice", "Double"
Me.Add "ÅÌﬂ", "«“ÅÌﬂ", " «ÅÌﬂ", , "Payk", "Integer"
Me.Add "”—Ì«·", "«“”—Ì«·", " «”—Ì«·", , "SerialNo", "Double"
Me.Add "ê—ÊÂ «’·Ì", "«“ê—ÊÂ «’·Ì", " «ê—ÊÂ «’·Ì", , "L1", "String"
Me.Add "ê—ÊÂ ›—⁄Ì", "«“ê—ÊÂ ›—⁄Ì", " «ê—ÊÂ ›—⁄Ì", , "L2", "String"
Me.Add "‰«„ ﬂ«·«", "«“‰«„ ﬂ«·«", " «‰«„ ﬂ«·«", , "Good", "String"
Me.Add "‘Ì› ", "«“‘Ì› ", " «‘Ì› ", , "Shift", "String"
Me.Add "ê«—”Ê‰", "«“ ê«—”Ê‰", " « ê«—”Ê‰", , "garson", "Integer"
Me.Add "ÅÌò", "«“ ÅÌò", " « ÅÌò", , "carrier", "Integer"
Me.Add "‘„«—Â ›«ò Ê—", "‘„«—Â ›«ò Ê—", "‘„«—Â ›«ò Ê—", , "FactorNo", "Integer"
Me.Add "›—„  ç«Å", "›—„  ç«Å", "›—„  ç«Å", , "PrintFormat", "Integer"
Me.Add "«‰»«—", "«“ «‰»«—", " « «‰»«—", , "Description", "String"
Me.Add "«‰»«—", "«“ ", " « ", , "StoreDescription", "String"
Me.Add "ê—œ‘", "«“", " « ", , "CalculatedGood", "String"
Me.Add "ÿ—› Õ”«»", "«“", " « ", , "CustomerType", "String"
Me.Add "Ã“∆Ì« ", "«“", " « ", , "Details", "String"
Me.Add "”«Ì“ ﬂ«€–", "«“", " « ", , "PaperType", "String"
Me.Add "”«· „«·Ì", "”«· „«·Ì", " ", , "AccountYear", "Integer"
Me.Add "‘⁄»Â", "‘⁄»Â", " ", , "Branch", "String"
Me.Add "‰Ê⁄", "‰Ê⁄", " ", , "Status", "String"
Me.Add "„Ì“", "«“ „Ì“", " « „Ì“ ", , "Table", "Integer"
Me.Add " — Ì»", "«“", " « ", , "SortOrder", "String"
Me.Add "„«Â", "«“ „«Â", " « „«Â ", , "Month", "Integer"
Me.Add "›—Ê‘‰œÂ", "«“ ›—Ê‘‰œÂ", " « ›—Ê‘‰œÂ", , "Seller", "Integer"
Me.Add "¬—«Ì‘ê—", "«“ ¬—«Ì‘ê—", " « ¬—«Ì‘ê—", , "HairDresser", "Integer"
Me.Add "‰Ê⁄ ÅÊ—”«‰ ", "«“", " « ", , "SellerReportType", "String"
Me.Add "Ê÷⁄Ì  ”›«—‘", "«“", " « ", , "OrderKind", "String"
Me.Add "‰ÕÊÂ ‰„«Ì‘", "«“", " « ", , "ShowKind", "String"
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    On Error Resume Next
    Set mvarRc = Nothing
    Set ClsGl = Nothing
    Set ObjNewMember = Nothing
    Set mcol = Nothing
'    mdifrm.fwBtnCtrl.SetFocus
End Sub



