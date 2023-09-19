VERSION 5.00
Begin VB.UserControl bigButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   DrawMode        =   7  'Invert
   DrawStyle       =   5  'Transparent
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
   ScaleHeight     =   3165
   ScaleWidth      =   3030
   Begin VB.Image imgNormal 
      Height          =   570
      Left            =   0
      Picture         =   "bigbutton.ctx":0000
      Top             =   2040
      Width           =   765
   End
   Begin VB.Image imgActive 
      Height          =   570
      Left            =   0
      Picture         =   "bigbutton.ctx":0725
      Top             =   630
      Width           =   765
   End
   Begin VB.Image imgPlaying 
      Height          =   210
      Left            =   0
      Top             =   2730
      Width           =   615
   End
   Begin VB.Image imgRecord 
      Height          =   210
      Left            =   0
      Top             =   1770
      Width           =   615
   End
   Begin VB.Image imgStop 
      Height          =   210
      Left            =   0
      Top             =   1500
      Width           =   615
   End
   Begin VB.Image imgPlay 
      Height          =   210
      Left            =   0
      Top             =   1230
      Width           =   615
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   195
      Left            =   990
      TabIndex        =   0
      Top             =   180
      Width           =   555
   End
   Begin VB.Image imgOff 
      Height          =   570
      Left            =   0
      Picture         =   "bigbutton.ctx":0E0C
      Top             =   0
      Width           =   765
   End
   Begin VB.Image imgON 
      Height          =   210
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "bigButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Just a small button I made :)

Enum bigModes
    mACTIVE
    mNORMAL
End Enum

Dim ButtonMode As bigModes
Public buttonstat As Boolean

Public Event Click()
Public Event ButtonDown()
Public Event ButtonUp()

Public Event MouseOver()

Public Sub SetButtonMode(nMode As bigModes)
    Select Case nMode
        Case Is = mACTIVE
            imgOff.Picture = imgActive.Picture
        Case Is = mNORMAL
            imgOff.Picture = imgNormal.Picture

    End Select
End Sub

Private Sub imgOff_Click()
    RaiseEvent Click
End Sub

Private Sub imgOff_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgOff.BorderStyle = 1
    buttonstat = Not buttonstat
End Sub

Private Sub imgOff_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgOff.BorderStyle = 0
End Sub


Private Sub UserControl_Paint()
    SetButtonMode ButtonMode
End Sub

Private Sub UserControl_Resize()
    lblCaption.Left = imgOff.Width + 100
    UserControl.Height = imgOff.Height
    UserControl.Height = imgOff.Height
    UserControl.Width = imgOff.Width
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

Public Property Get MODES() As bigModes
    MODES = ButtonMode
End Property

Public Property Let MODES(ByVal sNewValue As bigModes)
    ButtonMode = sNewValue
    UserControl.PropertyChanged "MODES"
    SetButtonMode ButtonMode
End Property




