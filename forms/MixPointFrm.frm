VERSION 5.00
Begin VB.Form MixPointFrm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Whoosh MP3 - Mixpoint Editor v1.1"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   Icon            =   "MixPointFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "Vocal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   37
      Top             =   2040
      Width           =   870
   End
   Begin VB.TextBox txtvocalstart 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1185
      TabIndex        =   33
      Text            =   "0"
      Top             =   2580
      Width           =   1350
   End
   Begin VB.PictureBox picBgSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   915
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   32
      Top             =   4305
      Width           =   1920
   End
   Begin VB.TextBox txtNewEndMix 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6825
      TabIndex        =   24
      Text            =   "0"
      Top             =   2565
      Width           =   1350
   End
   Begin VB.TextBox txtNewStartMix 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4035
      TabIndex        =   23
      Text            =   "0"
      Top             =   2580
      Width           =   1350
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Outtro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1140
      TabIndex        =   18
      Top             =   2040
      Width           =   900
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Intro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Value           =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmdSaveMixPoint 
      Caption         =   "Save new mixpoint/vocal start"
      Height          =   330
      Left            =   5325
      TabIndex        =   16
      Top             =   3840
      Width           =   2490
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      Height          =   330
      Left            =   8655
      TabIndex        =   14
      Top             =   1950
      Width           =   450
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "-"
      Height          =   330
      Left            =   8190
      TabIndex        =   13
      Top             =   1950
      Width           =   450
   End
   Begin VB.Frame Frame1 
      Caption         =   "Song Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   60
      TabIndex        =   8
      Top             =   2895
      Width           =   5190
      Begin VB.TextBox lblArtist 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   990
         TabIndex        =   31
         Top             =   615
         Width           =   4080
      End
      Begin VB.TextBox lblTitle 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   990
         TabIndex        =   30
         Top             =   300
         Width           =   4080
      End
      Begin VB.Label lblCategory 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2640
         TabIndex        =   28
         Top             =   960
         Width           =   2445
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1830
         TabIndex        =   27
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lblBPM 
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   12
         Top             =   990
         Width           =   495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "BPM:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   480
         TabIndex        =   11
         Top             =   990
         Width           =   510
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Artist:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   390
         TabIndex        =   10
         Top             =   585
         Width           =   510
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   495
         TabIndex        =   9
         Top             =   315
         Width           =   390
      End
   End
   Begin VB.TextBox txtInPos 
      BackColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   6315
      TabIndex        =   6
      Text            =   "0"
      Top             =   1950
      Width           =   1350
   End
   Begin VB.Timer InTimer 
      Enabled         =   0   'False
      Left            =   -225
      Top             =   3855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   345
      Left            =   8235
      TabIndex        =   5
      Top             =   3810
      Width           =   1140
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   330
      Left            =   5085
      TabIndex        =   4
      Top             =   1950
      Width           =   855
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   330
      Left            =   4215
      TabIndex        =   3
      Top             =   1950
      Width           =   855
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   330
      Left            =   3360
      TabIndex        =   2
      Top             =   1950
      Width           =   840
   End
   Begin VB.PictureBox pcIntroWave 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      FillColor       =   &H00FFFF80&
      ForeColor       =   &H0000FF00&
      Height          =   1665
      Left            =   30
      ScaleHeight     =   1605
      ScaleWidth      =   9300
      TabIndex        =   0
      Top             =   255
      Width           =   9360
      Begin VB.Line InLine 
         BorderColor     =   &H00FFFF00&
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   1560
      End
   End
   Begin VB.Label lblStartVocal 
      BackStyle       =   0  'Transparent
      Caption         =   "Startvocal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2610
      TabIndex        =   36
      Top             =   2625
      Width           =   1275
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Start vocal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2610
      TabIndex        =   35
      Top             =   2400
      Width           =   930
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "New Vocal Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1170
      TabIndex        =   34
      Top             =   2370
      Width           =   1380
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   $"MixPointFrm.frx":0442
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   765
      Left            =   5400
      TabIndex        =   29
      Top             =   2985
      Width           =   3960
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "New start mixpoint"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4020
      TabIndex        =   26
      Top             =   2370
      Width           =   1380
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "New end mixpoint"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   6825
      TabIndex        =   25
      Top             =   2355
      Width           =   1380
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Start mix:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5430
      TabIndex        =   22
      Top             =   2385
      Width           =   720
   End
   Begin VB.Label lblStartMix 
      BackStyle       =   0  'Transparent
      Caption         =   "Startmix:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5430
      TabIndex        =   21
      Top             =   2610
      Width           =   1275
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "End mix:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   8220
      TabIndex        =   20
      Top             =   2385
      Width           =   630
   End
   Begin VB.Label lblEndMix 
      BackStyle       =   0  'Transparent
      Caption         =   "Endmix:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   8220
      TabIndex        =   19
      Top             =   2610
      Width           =   1155
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom"
      Height          =   225
      Left            =   7725
      TabIndex        =   15
      Top             =   2010
      Width           =   405
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos"
      Height          =   225
      Left            =   5970
      TabIndex        =   7
      Top             =   2010
      Width           =   405
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Music data information"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   1695
   End
End
Attribute VB_Name = "MixPointFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prvBUFF(256) As Single
Dim IntroChan As Long
Dim OutroChan As Long
Dim PrevChan As Long
Public SamLen As Long
Dim OutPos As Long
Dim PrevChanel As MUSICFILE
Dim InInterval As Integer
Dim InSamLen As Long
Dim InStreamPos As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Dim rsMusic As ADODB.Recordset
Public idxmp As Long

Private Sub PaintBackground(dfrm As Form, BgSrc As PictureBox)
Dim x As Integer
Dim y As Integer
If FileExist(App.Path & "\bg.jpg") = False Then Exit Sub
BgSrc.Picture = LoadPicture("bg.jpg")
For y = 0 To dfrm.ScaleHeight Step BgSrc.ScaleHeight
    For x = 0 To dfrm.ScaleWidth Step BgSrc.ScaleWidth
         BitBlt dfrm.hdc, x, y, BgSrc.ScaleWidth, BgSrc.ScaleHeight, BgSrc.hdc, 0, 0, vbSrcCopy
    Next x
Next y
End Sub

Private Sub cmdClose_Click()
    cmdStop_Click
    Unload Me
End Sub


Private Sub DrawWave2()
Dim o As Integer
Dim data As Single
Dim co As ColorConstants

Me.MousePointer = 11
  InSamLen = pcIntroWave.Width * (2048 / InInterval)

  'load the mP3 file
  BASS_StreamFree IntroChan
  BASS_StreamFree PrevChan
  
  PrevChan = BASS_StreamCreateFile(BASSFALSE, PrevChanel.file, 0, 0, 0)
  IntroChan = BASS_StreamCreateFile(BASSFALSE, PrevChanel.file, 0, 0, BASS_MUSIC_DECODE Or BASS_MP3_SETPOS)
  pcIntroWave.Cls
  x = 0
  
  Call BASS_ChannelSetPosition(IntroChan, BASS_ChannelGetLength(IntroChan))
  InStreamPos = BASS_ChannelGetLength(IntroChan) - InSamLen
  Call BASS_ChannelSetPosition(IntroChan, InStreamPos)
  
  Do While InStreamPos < BASS_ChannelGetLength(IntroChan)
        InStreamPos = BASS_ChannelGetPosition(IntroChan)
        BASS_ChannelGetData IntroChan, prvBUFF(0), BASS_DATA_FFT512
        If InStreamPos < PrevChanel.SongEnd Then
            co = vbGreen
        End If
        If InStreamPos > PrevChanel.SongEnd Then
            co = vbRed
        End If
        
        For i = 1 To 5
            pcIntroWave.Line (x, (pcIntroWave.Height - 100) - (prvBUFF(i) * 2000))-(x, pcIntroWave.Height), co
        Next i
        
        x = x + InInterval
        DoEvents  'do other events
  Loop

txtInPos.text = BASS_ChannelGetLength(IntroChan) - InSamLen
Me.MousePointer = 1
End Sub


Private Sub DrawWave(offset As Long)
Dim o As Integer
Dim data As Single
Dim co As ColorConstants
Dim mLen As Long
Dim x As Integer
  
  InSamLen = pcIntroWave.Width * (2048 / InInterval)

  'load the mP3 file
  BASS_StreamFree IntroChan
  BASS_StreamFree PrevChan
  
  PrevChan = BASS_StreamCreateFile(BASSFALSE, PrevChanel.file, 0, 0, 0)
  IntroChan = BASS_StreamCreateFile(BASSFALSE, PrevChanel.file, 0, 0, BASS_MUSIC_DECODE Or BASS_MP3_SETPOS)
  pcIntroWave.Cls
  x = 0
  InStreamPos = 0
  
  mLen = BASS_ChannelGetLength(IntroChan)
  
  If InSamLen > mLen Then InSamLen = mLen
  
  Call BASS_ChannelSetPosition(IntroChan, offset)
  
  Do While InStreamPos < InSamLen
        InStreamPos = BASS_ChannelGetPosition(IntroChan)
        BASS_ChannelGetData IntroChan, prvBUFF(0), BASS_DATA_FFT512
        If InStreamPos < PrevChanel.SongStart Then
            co = vbRed
        End If
        If InStreamPos > PrevChanel.SongStart Then
            co = vbGreen
        End If
        For i = 1 To 3
            pcIntroWave.Line (x, (pcIntroWave.Height - 100) - (prvBUFF(i) * 2000))-(x, pcIntroWave.Height), co
        Next i
        x = x + InInterval
        DoEvents  'do other events
        
  Loop
End Sub

Private Sub DrawLine(x As Long, y As Long, x1 As Long, y1 As Long)
    Dim pt As POINTAPI
    pt.x = x
    pt.y = y
    
    MoveToEx pcIntroWave.hdc, x, y, pt
    LineTo pcIntroWave.hdc, x1, y1
    
End Sub


Private Sub DrawWave3(offset As Long)
Dim o As Integer
Dim data As Single
Dim co As ColorConstants
Dim mLen As Long
  
  InSamLen = pcIntroWave.Width * (2048 / InInterval)

  'load the mP3 file
  BASS_StreamFree IntroChan
  BASS_StreamFree PrevChan
  
  PrevChan = BASS_StreamCreateFile(BASSFALSE, PrevChanel.file, 0, 0, 0)
  IntroChan = BASS_StreamCreateFile(BASSFALSE, PrevChanel.file, 0, 0, BASS_MUSIC_DECODE Or BASS_MP3_SETPOS)
  pcIntroWave.Cls
  x = 0
  InStreamPos = 0
  
  mLen = BASS_ChannelGetLength(IntroChan)
  
  If InSamLen > mLen Then InSamLen = mLen
  
  Call BASS_ChannelSetPosition(IntroChan, offset)
  
  Do While InStreamPos < InSamLen
        InStreamPos = BASS_ChannelGetPosition(IntroChan)
        BASS_ChannelGetData IntroChan, prvBUFF(0), BASS_DATA_FFT512

        If InStreamPos < PrevChanel.vocalstart Then
            co = &H80FF&
        End If
        If InStreamPos > PrevChanel.vocalstart Then
            co = vbGreen
        End If
        For i = 1 To 3
            pcIntroWave.Line (x, (pcIntroWave.Height - 100) - (prvBUFF(i) * 2000))-(x, pcIntroWave.Height), co
        Next i
        x = x + InInterval
        DoEvents  'do other events
        
  Loop
End Sub


Private Sub cmdMinus_Click()
    InInterval = InInterval - 5
    If InInterval < 5 Then InInterval = 5
    Call cmdScan_Click
End Sub

Private Sub cmdPlay_Click()
        Call BASS_ChannelPlay(PrevChan, BASSTRUE, 0)
        Call BASS_ChannelSetPosition(PrevChan, txtInPos.text)
    
    InTimer.Interval = 1
    InTimer.Enabled = True
End Sub




Private Sub cmdPlus_Click()
    InInterval = InInterval + 5
    If InInterval > 20 Then InInterval = 20
    Call cmdScan_Click
End Sub

Private Sub cmdSaveMixPoint_Click()
    Dim sqltxt As String

    PrevChanel.SongStart = txtNewStartMix.text
    PrevChanel.SongEnd = txtNewEndMix.text
    PrevChanel.vocalstart = txtvocalstart.text
    
   
    sqltxt = "UPDATE music SET songstart=" & PrevChanel.SongStart & ",songend=" & PrevChanel.SongEnd & ",vocalstart=" & PrevChanel.vocalstart & " WHERE index=" & PrevChanel.Index
    
    ExecuteQuery sqltxt
    
    cmdScan_Click
    
End Sub

Private Sub cmdScan_Click()
If Option1.value = True Then
    DrawWave 0
End If
If Option2.value = True Then
    DrawWave2
End If

If Option3.value = True Then
    DrawWave3 0
End If
End Sub


Private Sub cmdStop_Click()
    Call BASS_ChannelStop(PrevChan)
    InTimer.Enabled = False
    InTimer.Interval = 0
End Sub


Private Sub Form_Load()
Dim idx As Integer
ChDrive App.Path
ChDir App.Path
Call ReadPlayList
  
  
  
InInterval = 10


    idxmp = MainFrm.idxmixpoint


    Set rsMusic = OpenRS("SELECT * FROM music WHERE index=" & idxmp & "")
    If rsMusic.RecordCount <> 0 Then
        PrevChanel.Index = rsMusic.Fields!Index
        PrevChanel.file = rsMusic.Fields!file
        PrevChanel.Title = rsMusic.Fields!Title
        PrevChanel.Artist = rsMusic.Fields!Artist
        PrevChanel.bpm = rsMusic.Fields!bpm
        PrevChanel.SongStart = rsMusic.Fields!SongStart
        PrevChanel.SongEnd = rsMusic.Fields!SongEnd
        PrevChanel.Category = rsMusic.Fields!Category
        PrevChanel.vocalstart = rsMusic.Fields!vocalstart
        
    
        lblTitle.text = PrevChanel.Title
        lblStartMix.Caption = PrevChanel.SongStart
        txtNewStartMix.text = PrevChanel.SongStart
        lblEndMix.Caption = PrevChanel.SongEnd
        txtNewEndMix.text = PrevChanel.SongEnd
        
        lblStartVocal.Caption = PrevChanel.vocalstart
        txtvocalstart.text = PrevChanel.vocalstart
        
        lblCategory.Caption = PrevChanel.Category
        lblBPM.Caption = PrevChanel.bpm
        
        cmdScan_Click
    End If



End Sub

Private Sub InTimer_Timer()
    txtInPos.text = BASS_ChannelGetPosition(PrevChan)
    
    If Option1.value = True Then
        InLine.x1 = (txtInPos.text / 2048) * InInterval
        InLine.X2 = InLine.x1
        If txtInPos.text > InSamLen Then cmdStop_Click
    End If

    If Option2.value = True Then
        InLine.x1 = ((txtInPos.text - (BASS_ChannelGetLength(IntroChan) - InSamLen)) / 2048) * InInterval
        InLine.X2 = InLine.x1
    End If


    If Option3.value = True Then
        InLine.x1 = (txtInPos.text / 2048) * InInterval
        InLine.X2 = InLine.x1
        If txtInPos.text > InSamLen Then cmdStop_Click
    End If

End Sub

Private Sub pcIntroWave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    If Option1.value = True Then
        txtInPos.text = (x * 2048) / InInterval
        txtNewStartMix.text = txtInPos.text
        InLine.x1 = x '(txtInPos.Text / 2048) * InInterval
        InLine.X2 = InLine.x1
    End If

    If Option2.value = True Then
        txtInPos.text = (BASS_ChannelGetLength(IntroChan) - InSamLen) + (x * 2048) / InInterval
        txtNewEndMix.text = txtInPos.text
        InLine.x1 = x
        InLine.X2 = InLine.x1
    End If
    
    If Option3.value = True Then
        txtInPos.text = (x * 2048) / InInterval
        txtvocalstart.text = txtInPos.text
        InLine.x1 = x '(txtInPos.Text / 2048) * InInterval
        InLine.X2 = InLine.x1
    End If
End If
End Sub


