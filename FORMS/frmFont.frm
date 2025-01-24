VERSION 5.00
Object = "{E6BA1CE2-2668-11D4-93D6-400100005168}#2.3#0"; "FAST2011.ocx"
Begin VB.Form frmFont 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "             ⁄ÌÌ‰ ›Ê‰  "
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3210
   Icon            =   "frmFont.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3210
   StartUpPosition =   2  'CenterScreen
   Begin FLWCtrls2.FWComboFont FWComboFont1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 '   FWComboFont1.Font
End Sub

Private Sub FWComboFont1_Change()
''''Dim obj As Object
''''Dim ObjectType As TypeObject
''''Dim varForm As Form
''''Dim frmAct As Form
''''
''''For Each varForm In Forms
''''    If VarActForm = varForm.Name Then
''''        Set frmAct = varForm
''''        Exit For
''''    End If
''''Next
''''    For Each obj In frmAct
''''
''''        If obj Is vbtextbox Then
''''            obj.FontName = "times new roman"
'''''                    Obj.Alignment = vbLeftJustify
''''        End If
''''    Next obj
        SaveSetting strMainKey, VarActForm, "Flexgrid_Name", FWComboFont1.Font.Name
        SaveSetting strMainKey, VarActForm, "Flexgrid_Size", FWComboFont1.Font.Size
        SaveSetting strMainKey, VarActForm, "Flexgrid_Bold", FWComboFont1.Font.Bold
  
End Sub
