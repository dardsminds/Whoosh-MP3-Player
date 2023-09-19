VERSION 5.00
Begin VB.UserControl Equalizer 
   BackColor       =   &H0000FFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
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
   ScaleHeight     =   1095
   ScaleWidth      =   210
   ToolboxBitmap   =   "Equalizer.ctx":0000
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   13
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   13
      Top             =   30
      Width           =   180
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   12
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   12
      Top             =   105
      Width           =   180
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   11
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   11
      Top             =   180
      Width           =   180
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   10
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   10
      Top             =   255
      Width           =   180
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   9
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   9
      Top             =   330
      Width           =   180
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   8
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   8
      Top             =   405
      Width           =   180
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   7
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   7
      Top             =   480
      Width           =   180
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   6
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   6
      Top             =   555
      Width           =   180
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   5
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   5
      Top             =   630
      Width           =   180
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   4
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   4
      Top             =   705
      Width           =   180
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   3
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   3
      Top             =   780
      Width           =   180
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   2
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   2
      Top             =   855
      Width           =   180
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   1
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   1
      Top             =   930
      Width           =   180
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   0
      Left            =   15
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   1005
      Width           =   180
   End
End
Attribute VB_Name = "Equalizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ValueChange(newValue As Integer)
Public Event FrequencyChange(newValue As Long)


Private mValue As Integer
Private mLedOffColor As OLE_COLOR
Private mLedOnColor As OLE_COLOR
Private mFrequency As Long


Public Property Get value() As Integer
    value = mValue
End Property

Public Property Let value(ByVal sNewValue As Integer)
    mValue = sNewValue
    UserControl.PropertyChanged "Value"
    RaiseEvent ValueChange(mValue)
    RefreshEqualizer
End Property


Public Property Get OffLedColor() As OLE_COLOR
    OffLedColor = mLedOffColor
End Property
Public Property Let OffLedColor(ByVal New_ForeColor As OLE_COLOR)
    mLedOffColor = New_ForeColor
    PropertyChanged "OffLedColor"
    RefreshEqualizer
End Property

Public Property Get LedOnColor() As OLE_COLOR
    LedOnColor = mLedOnColor
End Property
Public Property Let LedOnColor(ByVal New_ForeColor As OLE_COLOR)
    mLedOnColor = New_ForeColor
    PropertyChanged "LedOnColor"
    RefreshEqualizer
End Property

Private Sub picLed_Click(Index As Integer)
    mValue = Index
    UserControl.PropertyChanged "Value"
    RaiseEvent ValueChange(mValue)
    RefreshEqualizer
End Sub

Public Property Get Frequency() As Long
    Frequency = mFrequency
End Property
Public Property Let Frequency(ByVal New_Freq As Long)
    mFrequency = New_Freq
    PropertyChanged "Frequency"
End Property


Private Sub picLed_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    mValue = Index
    UserControl.PropertyChanged "Value"
    RaiseEvent ValueChange(mValue)
    RefreshEqualizer
End If
End Sub

Private Sub UserControl_Show()
    RefreshEqualizer
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Frequency", mFrequency, 1000
    .WriteProperty "Value", mValue, 0
    .WriteProperty "OffLedColor", mLedOffColor, picLed(0).BackColor
    .WriteProperty "LedOnColor", mLedOnColor, vbYellow
    
End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
   mFrequency = .ReadProperty("Frequency", 1000)
   mValue = .ReadProperty("Value", mValue)
   mLedOffColor = .ReadProperty("OffLedColor", picLed(0).BackColor)
   mLedOnColor = .ReadProperty("LedOnColor", vbYellow)
End With
End Sub

Private Sub RefreshEqualizer()
    Dim i As Integer
    For i = 0 To picLed.UBound
        picLed(i).BackColor = mLedOffColor
        picLed(i).BorderStyle = 0
        picLed(i).Height = 67
    Next i
    picLed(mValue).BackColor = mLedOnColor
        picLed(mValue).BorderStyle = 1
        picLed(mValue).Height = 68
    
End Sub
