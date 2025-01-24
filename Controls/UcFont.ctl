VERSION 5.00
Object = "{E6BA1CE2-2668-11D4-93D6-400100005168}#2.3#0"; "FAST2011.ocx"
Begin VB.UserControl UcFont 
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   ScaleHeight     =   675
   ScaleWidth      =   2535
   Begin FLWCtrls2.FWComboFont FWComboFont1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   " ⁄ÌÌ‰ ›Ê‰  Ê ”«Ì“"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "UcFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Dim m_FontName As String
Dim m_FontSize As Integer
Dim m_FontBold As String
Dim m_varActForm As String
Event FontProperty(m_FontName, m_FontSize, m_FontBold)
Public Property Get FontName() As String
    FontName = FWComboFont1.Font.Name
End Property
Public Property Let FontName(ByVal vData As String)
    FWComboFont1.Font.Name = vData
    m_FontName = vData
End Property
Public Property Get FontSize() As Integer
    FontSize = FWComboFont1.Font.Size
End Property
Public Property Let FontSize(ByVal vData As Integer)
    FWComboFont1.Font.Size = vData
    m_FontSize = vData
End Property
Public Property Get FontBold() As String
    FontBold = FWComboFont1.Font.Bold
End Property
Public Property Let FontBold(ByVal vData As String)
    FWComboFont1.Font.Bold = vData
    m_FontBold = vData
End Property
Public Property Get VarActForm() As String
    VarActForm = m_varActForm
End Property
Public Property Let VarActForm(ByVal vData As String)
    m_varActForm = vData
End Property
    
Private Sub FWComboFont1_Change()
    m_FontName = FWComboFont1.Font.Name
    m_FontSize = FWComboFont1.Font.Size
    m_FontBold = FWComboFont1.Font.Bold
    
    SaveSetting strMainKey, m_varActForm, "Flexgrid_Name", m_FontName
    SaveSetting strMainKey, m_varActForm, "Flexgrid_Size", m_FontSize
    SaveSetting strMainKey, m_varActForm, "Flexgrid_Bold", m_FontBold
    
    RaiseEvent FontProperty(m_FontName, m_FontSize, m_FontBold)

End Sub
