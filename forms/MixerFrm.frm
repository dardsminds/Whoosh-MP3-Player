VERSION 5.00
Object = "{34F681D0-3640-11CF-9294-00AA00B8A733}#1.0#0"; "DANIM.DLL"
Begin VB.Form MixerFrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00BE9B96&
   BorderStyle     =   0  'None
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "MixerFrm.frx":0000
   ScaleHeight     =   3585
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1905
      ScaleHeight     =   195
      ScaleWidth      =   465
      TabIndex        =   1
      Top             =   135
      Width           =   495
      Begin VB.Label lblCPU 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00.0"
         ForeColor       =   &H00C0C000&
         Height          =   225
         Left            =   75
         TabIndex        =   2
         Top             =   0
         Width           =   375
      End
   End
   Begin Whoosh.cpvSlider CrossFader 
      Height          =   225
      Left            =   240
      Top             =   2130
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   397
      BackColor       =   16777215
      SliderIcon      =   "MixerFrm.frx":14CFA
      Orientation     =   0
      RailPicture     =   "MixerFrm.frx":15094
      RailStyle       =   99
      Max             =   100
      Value           =   50
   End
   Begin VB.Timer MixTimer 
      Enabled         =   0   'False
      Left            =   240
      Top             =   2625
   End
   Begin VB.PictureBox BgSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   1200
      Picture         =   "MixerFrm.frx":1681A
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   2640
      Width           =   900
   End
   Begin Whoosh.GurhanButton cmdloop 
      Height          =   375
      Index           =   0
      Left            =   30
      TabIndex        =   4
      ToolTipText     =   "loop"
      Top             =   450
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      Caption         =   "1"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12030773
      ForeColor       =   4210688
   End
   Begin Whoosh.GurhanButton cmdloop 
      Height          =   390
      Index           =   1
      Left            =   30
      TabIndex        =   5
      ToolTipText     =   "loop"
      Top             =   840
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   688
      Caption         =   "2"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12030773
      ForeColor       =   4210688
   End
   Begin Whoosh.GurhanButton cmdloop 
      Height          =   390
      Index           =   2
      Left            =   30
      TabIndex        =   6
      ToolTipText     =   "loop"
      Top             =   1245
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   688
      Caption         =   "3"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12030773
      ForeColor       =   4210688
   End
   Begin Whoosh.GurhanButton cmdloop 
      Height          =   375
      Index           =   3
      Left            =   30
      TabIndex        =   7
      ToolTipText     =   "loop"
      Top             =   1650
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      Caption         =   "4"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12030773
      ForeColor       =   4210688
   End
   Begin Whoosh.GurhanButton sam 
      Height          =   390
      Index           =   0
      Left            =   2250
      TabIndex        =   8
      ToolTipText     =   "sample"
      Top             =   435
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   688
      Caption         =   "1"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12030773
      ForeColor       =   4210688
   End
   Begin Whoosh.GurhanButton sam 
      Height          =   360
      Index           =   1
      Left            =   2250
      TabIndex        =   9
      ToolTipText     =   "sample"
      Top             =   840
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   635
      Caption         =   "2"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12030773
      ForeColor       =   4210688
   End
   Begin Whoosh.GurhanButton sam 
      Height          =   390
      Index           =   2
      Left            =   2250
      TabIndex        =   10
      ToolTipText     =   "sample"
      Top             =   1635
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   688
      Caption         =   "3"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12030773
      ForeColor       =   4210688
   End
   Begin Whoosh.GurhanButton sam 
      Height          =   405
      Index           =   3
      Left            =   2250
      TabIndex        =   11
      ToolTipText     =   "sample"
      Top             =   1215
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   714
      Caption         =   "4"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12030773
      ForeColor       =   4210688
   End
   Begin Whoosh.cpvSlider sldrLoopPitch 
      Height          =   150
      Left            =   840
      ToolTipText     =   "Loop speed"
      Top             =   150
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   265
      BackColor       =   8195858
      SliderIcon      =   "MixerFrm.frx":16B49
      Orientation     =   0
      RailPicture     =   "MixerFrm.frx":16DF3
      RailStyle       =   99
      ShowValueTip    =   0   'False
      Min             =   50
      Max             =   150
      Value           =   100
   End
   Begin DirectAnimationCtl.DAViewerControl DAControl 
      Height          =   1515
      Left            =   465
      TabIndex        =   12
      Top             =   480
      Width           =   1545
      OpaqueForHitDetect=   -1  'True
      UpdateInterval  =   0.033
   End
   Begin VB.Label lblAutoDJ 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto DJ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   105
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00B79335&
      BorderWidth     =   3
      Height          =   285
      Left            =   210
      Shape           =   4  'Rounded Rectangle
      Top             =   2085
      Width           =   2025
   End
End
Attribute VB_Name = "MixerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MixCount As Integer
Public CurrentPlayer As Integer
Dim MFade(100) As Integer
Dim MFadeB(100) As Integer
Dim L(4) As Boolean
Public Mixing As Boolean
Dim LoopChan As Long
Dim PFile As String
Const maxp = 1.2
Const minp = 0.8
Dim freq As Long

Public Function MixNow()
If Mp3(1).Estado = IDLE Then Player1.LoadMP3
If Mp3(2).Estado = IDLE Then Player2.LoadMP3
Mix = True
End Function

Private Sub cmdloop_Click(Index As Integer)
Select Case Index
Case Is = 0
    L(Index) = Not L(Index)
    If L(Index) = True Then
        PFile = App.Path & "\loops\loop1.wav"
        If FileExist(PFile) = False Then Exit Sub
        cmdloop(Index).BackColor = &H80C0FF
    
        'kung naga tokar hay e stop
        If BASS_ChannelIsActive(LoopChan) = BASS_ACTIVE_PLAYING Then BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
        LoopChan = BASS_StreamCreateFile(BASSFALSE, PFile, 0, 0, BASS_SAMPLE_LOOP)
      'set pitch
        Call BASS_ChannelGetAttributes(LoopChan, freq, vbNull, vbNull)
        With sldrLoopPitch
        .max = (freq * maxp) / 4
        .Min = (freq * minp) / 4
        .value = freq / 4
        End With
        Call BASS_StreamPlay(LoopChan, 0, BASS_SAMPLE_LOOP)
    Else
        BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
        cmdloop(Index).BackColor = &HB08B73
    End If
Case Is = 1
    L(Index) = Not L(Index)
    If L(Index) = True Then
        PFile = App.Path & "\loops\loop2.wav"
        If FileExist(PFile) = False Then Exit Sub
        cmdloop(Index).BackColor = &H80C0FF
        
        'kung naga tokar hay e stop
        If BASS_ChannelIsActive(LoopChan) = BASS_ACTIVE_PLAYING Then BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
        LoopChan = BASS_StreamCreateFile(BASSFALSE, PFile, 0, 0, BASS_SAMPLE_LOOP)
      'set pitch
        Call BASS_ChannelGetAttributes(LoopChan, freq, vbNull, vbNull)
        With sldrLoopPitch
        .max = (freq * maxp) / 4
        .Min = (freq * minp) / 4
        .value = freq / 4
        End With
        Call BASS_StreamPlay(LoopChan, 0, BASS_SAMPLE_LOOP)
    Else
        BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
        cmdloop(Index).BackColor = &HB08B73
        
    End If
Case Is = 2
    L(Index) = Not L(Index)
    If L(Index) = True Then
        PFile = App.Path & "\loops\loop3.wav"
        If FileExist(PFile) = False Then Exit Sub
        cmdloop(Index).BackColor = &H80C0FF
        
        'kung naga tokar hay e stop
        If BASS_ChannelIsActive(LoopChan) = BASS_ACTIVE_PLAYING Then BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
        LoopChan = BASS_StreamCreateFile(BASSFALSE, PFile, 0, 0, BASS_SAMPLE_LOOP)
      'set pitch
        Call BASS_ChannelGetAttributes(LoopChan, freq, vbNull, vbNull)
        With sldrLoopPitch
        .max = (freq * maxp) / 4
        .Min = (freq * minp) / 4
        .value = freq / 4
        End With
        Call BASS_StreamPlay(LoopChan, 0, BASS_SAMPLE_LOOP)
    Else
        BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
        cmdloop(Index).BackColor = &HB08B73
        
    End If
Case Is = 3
    L(Index) = Not L(Index)
    If L(Index) = True Then
        PFile = App.Path & "\loops\loop4.wav"
        If FileExist(PFile) = False Then Exit Sub
        cmdloop(Index).BackColor = &H80C0FF
        
        'kung naga tokar hay e stop
        If BASS_ChannelIsActive(LoopChan) = BASS_ACTIVE_PLAYING Then BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
        LoopChan = BASS_StreamCreateFile(BASSFALSE, PFile, 0, 0, BASS_SAMPLE_LOOP)
      'set pitch
        Call BASS_ChannelGetAttributes(LoopChan, freq, vbNull, vbNull)
        With sldrLoopPitch
        .max = (freq * maxp) / 4
        .Min = (freq * minp) / 4
        .value = freq / 4
        End With
        Call BASS_StreamPlay(LoopChan, 0, BASS_SAMPLE_LOOP)
    Else
        BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
        cmdloop(Index).BackColor = &HB08B73
        
    End If
End Select
sldrLoopPitch_ValueChanged
End Sub

Public Sub CrossFader_ValueChanged()
Call BASS_ChannelSetAttributes(BASS_FX_TempoGetResampledHandle(Mp3(1).chan), -1, MFade(CrossFader.value), -101)
Call BASS_ChannelSetAttributes(BASS_FX_TempoGetResampledHandle(Mp3(2).chan), -1, MFadeB(CrossFader.value), -101)
End Sub

Private Sub Form_Load()
Dim freq As Long
'arrays of fader value A scroll
For a = 0 To 50
MFade(a) = 100
Next a
For a = 50 To 100
MFade(a) = 100 - B
B = B + 2
Next a

'arrays of fader value b scroll
For a = 0 To 50
MFadeB(a) = a * 2
Next a
For a = 50 To 100
MFadeB(a) = 100
Next a

CurrentPlayer = 0
'If FileExist(App.Path & "\mixer.bmp") = True Then
'Me.Picture = LoadPicture("mixer.bmp")
'End If
Me.Height = 2550
Me.ScaleWidth = 2420
'PaintBackground Me, BgSrc


  Set m = DAControl.PixelLibrary
  Set fillImg = m.ImportImage(App.Path & "\ttable.jpg")
  Set ovalImg = m.Oval(100, 100).Fill(m.DefaultLineStyle, fillImg)
  Set rotXf = m.Rotate3RateDegrees(m.Vector3(0, 0, 1), 120).ParallelTransform2().Inverse
  Set finalImg = ovalImg.Transform(rotXf)
  DAControl.Image = finalImg
  DAControl.BackgroundImage = m.SolidColorImage(m.Silver)

  DAControl.start
  
End Sub




Private Sub sam_Click(Index As Integer)
Select Case Index
Case Is = 0
    BASS_SamplePlayEx Sample1, 0, -1, 100, Int((201 * Rnd) - 100), BASSFALSE
Case Is = 1
    BASS_SamplePlayEx Sample2, 0, -1, 100, Int((201 * Rnd) - 100), BASSFALSE
Case Is = 2
    BASS_SamplePlayEx Sample3, 0, -1, 100, Int((201 * Rnd) - 100), BASSFALSE
Case Is = 3
    BASS_SamplePlayEx Sample4, 0, -1, 100, Int((201 * Rnd) - 100), BASSFALSE
End Select
End Sub

Private Sub sldrLoopPitch_ValueChanged()
    Call BASS_ChannelSetAttributes(LoopChan, CLng(sldrLoopPitch.value) * 4, -1, -101)
End Sub
Private Sub PaintBackground(dfrm As Form, BgSrc As PictureBox)
Dim X As Integer
Dim Y As Integer
For Y = 0 To dfrm.ScaleHeight Step BgSrc.ScaleHeight
    For X = 0 To dfrm.ScaleWidth Step BgSrc.ScaleWidth
         bitblt dfrm.hdc, X, Y, BgSrc.ScaleWidth, BgSrc.ScaleHeight, BgSrc.hdc, 0, 0, vbSrcCopy
         DoEvents
    Next X
Next Y
End Sub

