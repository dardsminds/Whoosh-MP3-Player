VERSION 5.00
Begin VB.UserControl dbutton 
   BackColor       =   &H00808080&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LockControls    =   -1  'True
   ScaleHeight     =   2355
   ScaleWidth      =   2010
   ToolboxBitmap   =   "dbutton.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   960
      Top             =   1740
   End
   Begin VB.Image imgPlaying 
      Height          =   210
      Left            =   0
      Picture         =   "dbutton.ctx":0312
      Top             =   1350
      Width           =   615
   End
   Begin VB.Image imgNormal 
      Height          =   210
      Left            =   0
      Picture         =   "dbutton.ctx":0A1E
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image imgRecord 
      Height          =   210
      Left            =   0
      Picture         =   "dbutton.ctx":6024
      Top             =   810
      Width           =   615
   End
   Begin VB.Image imgStop 
      Height          =   210
      Left            =   0
      Picture         =   "dbutton.ctx":B9F0
      Top             =   540
      Width           =   615
   End
   Begin VB.Image imgPlay 
      Height          =   210
      Left            =   0
      Picture         =   "dbutton.ctx":11248
      Top             =   270
      Width           =   615
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   195
      Left            =   660
      TabIndex        =   0
      Top             =   30
      Width           =   555
   End
   Begin VB.Image imgOff 
      Height          =   210
      Left            =   0
      Picture         =   "dbutton.ctx":16C78
      Top             =   0
      Width           =   615
   End
   Begin VB.Image imgON 
      Height          =   210
      Left            =   0
      Picture         =   "dbutton.ctx":1C27E
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "dbutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Just a small button I made :)

Enum bMODES
    mNORMAL
    mPLAY
    mSTOP
    mRECORD
    mPLAYING
End Enum

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private mpoiCursorPos As POINTAPI
Dim ButtonMode As bMODES
Dim flashing As Boolean

Public Event Click()
Public Event ButtonDown()
Public Event ButtonUp()

Public Event MouseOver()

Public Sub SetButtonMode(nMode As bMODES)
    Timer1.Enabled = False

    Select Case nMode
        Case Is = mNORMAL
            imgOff.Picture = imgNormal.Picture
        Case Is = mPLAY
            imgOff.Picture = imgPlay.Picture
        Case Is = mSTOP
            imgOff.Picture = imgStop.Picture
        Case Is = mRECORD
            imgOff.Picture = imgRecord.Picture
        Case Is = mPLAYING
            Timer1.Enabled = True
    End Select
End Sub

Private Sub imgOff_Click()
    RaiseEvent Click
End Sub

Private Sub imgOff_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgOff.BorderStyle = 1
End Sub

Private Sub imgOff_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgOff.BorderStyle = 0
End Sub

Private Sub Timer1_Timer()
    flashing = Not flashing
    If flashing = False Then
       imgOff.Picture = imgPlay.Picture
    Else
       imgOff.Picture = imgPlaying.Picture
    End If
    
End Sub

Private Sub UserControl_Paint()
    SetButtonMode ButtonMode
End Sub

Private Sub UserControl_Resize()
    lblCaption.Left = imgOff.Width + 100
    UserControl.Height = imgOff.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "MODES", ButtonMode, mNORMAL
End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    ButtonMode = .ReadProperty("MODES", mNORMAL)
End With
End Sub

Public Property Get MODES() As bMODES
    MODES = ButtonMode
End Property

Public Property Let MODES(ByVal sNewValue As bMODES)
    ButtonMode = sNewValue
    UserControl.PropertyChanged "MODES"
    SetButtonMode ButtonMode
End Property




