VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9225
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   8640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1024
      ImageHeight     =   700
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSplash.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Left            =   720
      Top             =   7200
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   7080
   End
   Begin Total.ucAniGIF ucAniGIF 
      Height          =   8355
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   11265
      _ExtentX        =   29395
      _ExtentY        =   22040
      GIF             =   "frmSplash.frx":B6FEF
      Delay           =   30
      Stretch         =   10
      Loops           =   1
      DelayLoad       =   0
   End
   Begin VB.Label LblSoftwareRegistrationNotice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "«Ì‰ ‰—„ «›“«— »« ⁄‰Ê«‰ ”Ì” „ „ﬂ«‰Ì“Â —” Ê—«‰Ì- ”„— 2 ° œ— ‘Ê—«Ì ⁄«·Ì «‰›Ê—„« Ìﬂ ﬂ‘Ê— »« ‘„«—Â 203370 »Â À»  —”ÌœÂ «” ."
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   8350
      Width           =   11250
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Keep these things in mind when you play.

' 1. VB is single threaded. The more animated GIFs on the form, the longer it
'   will take for all of them to begin animation.
'   - The DelayAnimation property helps the form load faster
'   - When DelayAnimation=False, form will take more time to initially display

' 2. Compile the usercontrol. If uncompiled, you can expect these common annoyances:
'   - When MsgBox pops up, images disappear until Msgbox is closed
'   - Dragging other windows over a static GIF (stopped animating) may not repaint immediately

' 3. The dotted/thick border in design time does not disappear. It is purposely painted to:
'   - Identify the animated GIF control from a normal VB image control
'   - Show you the bounds of the overall image

' 4. Do not overlap image controls if possible.
'   - Overlapped controls forward paint events to every control above it in the zOrder
'   - Having several controls overlapped theoretically can bog down your application

' 5. For performance reasons, you should pause or stop animation when your app is minimized

Private Sub Form_Activate()
'    Call objAnimateGIF.PlayGIF
    Timer1.Enabled = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 27  ' Esc
                    Timer1.Enabled = False
                    Timer2.Enabled = False
                    FrmLogin.Show
                    Sleep 500
                    Unload Me
              End Select

    End Select
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim i As Integer, Action As AnimationActions
    Select Case Index
        Case 0: Action = gfaStop
        Case 1: Action = gfaPause
        Case 2: Action = gfaForward
        Case 3: Action = gfaPlay
    End Select
    For i = ucAniGIF.LBound To ucAniGIF.UBound
        ucAniGIF(i).Action = Action
    Next
End Sub
Private Sub Form_Load()
'    chkMirror.BackColor = Me.BackColor
'    chkModPal.BackColor = Me.BackColor
'    chkModClrTbl2.BackColor = Me.BackColor
'    chkSolidBkg.BackColor = Me.BackColor
        
    Timer1.Interval = 10000
    Timer2.Interval = 8000
End Sub

Private Sub Form_Resize()
    ' Example of pausing animation while minimized
    Dim i As Integer
    Dim oldAction As AnimationActions, newAction As AnimationActions
    
    If Me.WindowState = vbMinimized Then
        oldAction = gfaPlay
        newAction = gfaPause
    Else
        oldAction = gfaPause
        newAction = gfaPlay
    End If
    
    For i = ucAniGIF.LBound To ucAniGIF.UBound
        If ucAniGIF(i).Action = oldAction Then ucAniGIF(i).Action = newAction
    Next

End Sub


Private Sub ucAniGIF_LoopsEnded(Index As Integer)
    ' should you want to know when a GIF terminates its
    ' loop and stops animating. You will also get this
    ' event for a single frame GIF each time it is displayed.
    ucAniGIF(Index).Action = gfaReset ' simply restart
End Sub

Private Sub ucAniGIF_RemoteLoadComplete(Index As Integer, ByVal gifWidth As Single, ByVal gifHeight As Single, ByRef Cancel As Boolean)
    ' When you called the LoadAnimatedGIF_Remote routine, (See Command2_Click)
    ' this event will be fired if the file was successfully read and the header
    ' of the file indicates it is a GIF.
    
    ' Set Cancel to True to prevent loading it.
    ' Otherwise, it will be displayed with the current settings of the usercontrol
    
    With ucAniGIF(Index)                ' to be a little thorough
        Set .AnimatedGIF = Nothing  ' remove previous image before changing attributes
        .Stretch = gfsShrinkScaleToFit ' set scale
        .DelayAnimation = gfdNone ' set delay mode
        .Mirrored = gfmNone         ' set mirror options
        .Enabled = True             ' enable it
    End With                        ' next, the image will be processed and displayed

End Sub

Private Sub ucAniGIF_RemoteLoadFailure(Index As Integer)
    ' When you called the LoadAnimatedGIF_Remote routine, this event will be fired if the
    ' file was NOT successfully read, errors occurred or the header of the file indicates
    ' it is NOT a GIF.
    MsgBox "Failed to download/read the remote GIF file. Possible server is down or " & vbCrLf & _
        "the GIF no longer exists. To test this functionality..." & vbCrLf & _
        "1. Go to any website that is displaying a GIF" & vbCrLf & _
        "2. Right click on the GIF and select Properties from the menu" & vbCrLf & _
        "3. Highlight the complete URL and copy it: Right click, copy" & vbCrLf & _
        "4. Paste the URL into the Command2_Click event. Try again.", vbInformation + vbOKOnly
'    Command2.Enabled = False
End Sub


Private Sub Timer1_Timer()
    Unload Me
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    FrmLogin.Show
End Sub
