VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ConfigFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Audio Configuration"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "ConfigFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Audio Quality"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   90
      TabIndex        =   11
      Top             =   2460
      Width           =   3285
      Begin VB.OptionButton optLowQ 
         Caption         =   "Low Quality (Pentium 1)"
         Height          =   285
         Left            =   390
         TabIndex        =   13
         Top             =   810
         Width           =   2205
      End
      Begin VB.OptionButton optHighQ 
         Caption         =   "High Quality (for fast CPU)"
         Height          =   315
         Left            =   390
         TabIndex        =   12
         Top             =   420
         Width           =   2625
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   3480
      Width           =   1125
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4740
      TabIndex        =   9
      Top             =   3480
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Caption         =   "Memory buffer length"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   90
      TabIndex        =   5
      Top             =   1470
      Width           =   5715
      Begin MSComctlLib.Slider Slider1 
         Height          =   405
         Left            =   300
         TabIndex        =   6
         Top             =   270
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   714
         _Version        =   393216
         Min             =   5
         Max             =   25
         SelStart        =   15
         TickFrequency   =   5
         Value           =   15
      End
      Begin VB.Label Label3 
         Caption         =   "2.5ms"
         Height          =   225
         Left            =   4560
         TabIndex        =   8
         Top             =   690
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "0.5ms"
         Height          =   225
         Left            =   450
         TabIndex        =   7
         Top             =   690
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Output Devices"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   5715
      Begin VB.CheckBox chkCue 
         Caption         =   "Enabled"
         Height          =   285
         Left            =   4650
         TabIndex        =   15
         Top             =   870
         Width           =   885
      End
      Begin VB.CheckBox chkOnAir 
         Caption         =   "Enabled"
         Height          =   285
         Left            =   4650
         TabIndex        =   14
         Top             =   390
         Width           =   885
      End
      Begin VB.ComboBox cmbCue 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   900
         Width           =   3195
      End
      Begin VB.ComboBox cmbOnAir 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   3195
      End
      Begin VB.Label Label2 
         Caption         =   "Cue Device"
         Height          =   225
         Left            =   300
         TabIndex        =   3
         Top             =   930
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "On Air Device"
         Height          =   225
         Left            =   150
         TabIndex        =   1
         Top             =   450
         Width           =   1035
      End
   End
End
Attribute VB_Name = "ConfigFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim freqRate As String
Dim memBuffer As String

Private Sub cmdApply_Click()
    memBuffer = Slider1.value / 10
    Call SaveINI("OUTPUT", "DeviceEnabled", chkOnAir.value)
    Call SaveINI("OUTPUT", "CueDeviceEnabled", chkCue.value)
    Call SaveINI("OUTPUT", "Device", cmbOnAir.ListIndex + 1)
    Call SaveINI("OUTPUT", "CueDevice", cmbCue.ListIndex + 1)
    Call SaveINI("OUTPUT", "Quality", freqRate)
    Call SaveINI("OUTPUT", "MemoryBuffer", memBuffer)
    MsgBox "Settings has been altered, you must restart the program to take effect", vbOKOnly, "Settings Updated"
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub



Private Sub Form_Load()
Dim text As String
Dim c As Integer

cmbOnAir.Clear
cmbCue.Clear
optHighQ.value = True
  
    c = 1
    While BASS_GetDeviceDescription(c)
        text = VBStrFromAnsiPtr(BASS_GetDeviceDescription(c))
        cmbOnAir.AddItem text
        cmbCue.AddItem text
        c = c + 1
    Wend
    
   If (cmbOnAir.ListCount > 1) Then
       cmbOnAir.ListIndex = AUDIO.OnAirDevice
       cmbCue.ListIndex = AUDIO.CueDevice
   Else
       cmbOnAir.ListIndex = 0
       cmbCue.ListIndex = 0
   End If
   
    chkOnAir.value = ReadINI("OUTPUT", "DeviceEnabled", "1")
    chkCue.value = ReadINI("OUTPUT", "CueDeviceEnabled", "0")

    Slider1.value = ReadINI("OUTPUT", "MemoryBuffer", "0.5") * 10
    If ReadINI("OUTPUT", "Quality", "44100") = 44100 Then
        optHighQ.value = True
    Else
        optLowQ.value = True
    End If
   
End Sub

Private Sub optHighQ_Click()
    freqRate = 44100
End Sub

Private Sub optLowQ_Click()
    freqRate = 22050
End Sub
