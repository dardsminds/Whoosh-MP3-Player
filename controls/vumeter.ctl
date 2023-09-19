VERSION 5.00
Begin VB.UserControl vumeter 
   BackColor       =   &H0000FFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   810
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
   ScaleHeight     =   1500
   ScaleWidth      =   810
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   300
      Top             =   960
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   19
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   19
      Top             =   0
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   18
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   18
      Top             =   75
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   17
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   17
      Top             =   150
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   16
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   16
      Top             =   225
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   15
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   15
      Top             =   300
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   14
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   14
      Top             =   375
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   13
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   13
      Top             =   450
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   12
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   12
      Top             =   525
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   11
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   11
      Top             =   600
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   10
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   10
      Top             =   675
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   9
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   9
      Top             =   750
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   8
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   8
      Top             =   825
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   7
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   7
      Top             =   900
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   6
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   6
      Top             =   975
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   5
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   5
      Top             =   1050
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   4
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   4
      Top             =   1125
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   3
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   2
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   2
      Top             =   1275
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   1
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   1
      Top             =   1350
      Width           =   120
   End
   Begin VB.PictureBox picLed 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   67
      Index           =   0
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   120
      TabIndex        =   0
      Top             =   1425
      Width           =   120
   End
End
Attribute VB_Name = "vumeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ValueChange(newValue As Long)
Public Event FrequencyChange(newValue As Long)


Private mValue As Long
Private mPeak As Integer
Private mLedOffColor As OLE_COLOR
Private mLedOnColor As OLE_COLOR
Private mFrequency As Long

Public Property Get value() As Long
    value = mValue
End Property

Public Property Let value(ByVal sNewValue As Long)
    mValue = Round(sNewValue)
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

Public Property Get Frequency() As Long
    Frequency = mFrequency
End Property
Public Property Let Frequency(ByVal New_Freq As Long)
    mFrequency = New_Freq
    PropertyChanged "Frequency"
End Property

Private Sub Timer1_Timer()
    mPeak = mPeak - 1
    If mPeak < 1 Then Timer1.Enabled = False
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 120
    UserControl.Height = 1500
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
    Dim a As Integer
    
    'Timer1.Enabled = True
    
    For i = 0 To picLed.UBound
        picLed(i).BackColor = mLedOffColor
        picLed(i).BorderStyle = 0
    Next i
    a = 0
    
   
    For i = 10 To 200 Step 10
        If mValue >= i Then
            picLed(a).BackColor = mLedOnColor
            'mPeak = a
        End If
        a = a + 1
    Next i
    
End Sub
