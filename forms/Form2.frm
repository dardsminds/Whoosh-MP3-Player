VERSION 5.00
Begin VB.Form Player2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   166
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8B479&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   300
      ScaleHeight     =   1605
      ScaleWidth      =   3810
      TabIndex        =   14
      Top             =   330
      Width           =   3810
      Begin VB.PictureBox analyzer 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8B479&
         BorderStyle     =   0  'None
         FillColor       =   &H00E8B479&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   990
         Left            =   45
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   174
         TabIndex        =   15
         Top             =   15
         Width           =   2610
      End
      Begin VB.Label Player_A 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8B479&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   33
         Top             =   1035
         Width           =   3120
      End
      Begin VB.Label lblBPM 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   2985
         TabIndex        =   32
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "BPM"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   105
         Index           =   6
         Left            =   2715
         TabIndex        =   31
         Top             =   885
         Width           =   195
      End
      Begin VB.Label Player_A 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   2985
         TabIndex        =   30
         Top             =   210
         Width           =   675
      End
      Begin VB.Label Player_A 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2985
         TabIndex        =   29
         Top             =   630
         Width           =   675
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "REM"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   105
         Index           =   2
         Left            =   2715
         TabIndex        =   28
         Top             =   690
         Width           =   195
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "TRIG"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   75
         Index           =   3
         Left            =   2655
         TabIndex        =   27
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "LEN"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   4
         Left            =   2715
         TabIndex        =   26
         Top             =   60
         Width           =   225
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CUR"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   105
         Index           =   5
         Left            =   2715
         TabIndex        =   25
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Player_A 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2985
         TabIndex        =   24
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Player_A 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2985
         TabIndex        =   23
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "TITLE"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   0
         Left            =   30
         TabIndex        =   22
         Top             =   1065
         Width           =   345
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ARTIST"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   1
         Left            =   15
         TabIndex        =   21
         Top             =   1290
         Width           =   435
      End
      Begin VB.Label Player_A 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   20
         Top             =   1245
         Width           =   3120
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ALBUM"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   7
         Left            =   30
         TabIndex        =   19
         Top             =   1500
         Width           =   435
      End
      Begin VB.Label lblAlbum 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   540
         TabIndex        =   18
         Top             =   1455
         Width           =   1755
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "GENRE"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   8
         Left            =   2325
         TabIndex        =   17
         Top             =   1530
         Width           =   360
      End
      Begin VB.Label lblGenre 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2655
         TabIndex        =   16
         Top             =   1455
         Width           =   1005
      End
   End
   Begin VB.CommandButton fx4 
      Caption         =   "H-pitch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3630
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2190
      Width           =   570
   End
   Begin VB.CommandButton fx3 
      Caption         =   "Echo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3045
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2190
      Width           =   570
   End
   Begin VB.CommandButton fx2 
      Caption         =   "Flanger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2190
      Width           =   570
   End
   Begin VB.CommandButton fx1 
      Caption         =   "Reverb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1875
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2190
      Width           =   570
   End
   Begin VB.CommandButton cmdResetBPM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4350
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1770
      Width           =   255
   End
   Begin VB.PictureBox gph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      DrawStyle       =   1  'Dash
      FillStyle       =   2  'Horizontal Line
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   2910
      ScaleHeight     =   1455
      ScaleWidth      =   180
      TabIndex        =   8
      Top             =   3060
      Width           =   210
   End
   Begin VB.PictureBox abuff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E8B479&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   120
      ScaleHeight     =   1365
      ScaleWidth      =   2685
      TabIndex        =   6
      Top             =   3075
      Width           =   2715
   End
   Begin VB.PictureBox bpmbuff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   1500
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   4
      Top             =   2745
      Width           =   135
   End
   Begin Whoosh.cpvSlider pitchSlider 
      Height          =   1380
      Left            =   4380
      Top             =   300
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   423
      BackColor       =   4210688
      SliderIcon      =   "Form2.frx":2775E
      RailPicture     =   "Form2.frx":279F0
      RailStyle       =   99
      Value           =   5
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   180
      Top             =   2700
   End
   Begin Whoosh.GurhanButton StopDeck 
      Height          =   345
      Left            =   1230
      TabIndex        =   1
      ToolTipText     =   "Stop current playing track"
      Top             =   2115
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      Caption         =   ""
      Picture         =   "Form2.frx":27A0C
      PictureWidth    =   29
      PictureHeight   =   23
      PictureSize     =   2
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
   End
   Begin Whoosh.GurhanButton DeckPlay 
      Height          =   345
      Left            =   780
      TabIndex        =   2
      ToolTipText     =   "Play loaded track"
      Top             =   2115
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      Caption         =   ""
      Picture         =   "Form2.frx":28246
      PictureWidth    =   29
      PictureHeight   =   23
      PictureSize     =   2
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
   End
   Begin Whoosh.GurhanButton DeckLoad 
      Height          =   345
      Left            =   330
      TabIndex        =   3
      ToolTipText     =   "Load next track from playlist"
      Top             =   2115
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      Caption         =   ""
      Picture         =   "Form2.frx":28A80
      PictureWidth    =   29
      PictureHeight   =   23
      PictureSize     =   2
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
   End
   Begin Whoosh.GurhanButton fx5 
      Height          =   150
      Left            =   4350
      TabIndex        =   5
      ToolTipText     =   "Dynamic gain "
      Top             =   2190
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   265
      Caption         =   ""
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Raised          =   -1  'True
      ForeColor       =   16711680
   End
   Begin Whoosh.cpvSlider pos 
      Height          =   120
      Left            =   330
      Top             =   1995
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   212
      BackColor       =   8421504
      SliderIcon      =   "Form2.frx":292BA
      Orientation     =   0
      RailPicture     =   "Form2.frx":2978C
      RailStyle       =   99
      ShowValueTip    =   0   'False
      Max             =   100
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   195
      Left            =   4320
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AGC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   150
      Left            =   4260
      TabIndex        =   7
      Top             =   1980
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   1440
      Left            =   4350
      Shape           =   4  'Rounded Rectangle
      Top             =   270
      Width           =   240
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "DECK B : IDLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E8B479&
      Height          =   165
      Left            =   1455
      TabIndex        =   0
      Top             =   30
      Width           =   1605
   End
End
Attribute VB_Name = "Player2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const maxp = 1.6
Const minp = 0.4
Dim freq As Long
Dim vol As Integer
Dim LoadPoint As Long

Dim i As Integer
Dim j As Integer
Dim file As String
Dim SampleTemp(0 To 300) As Single


Public MeReadyToStop As Boolean

Public Function StopMP3()
'    Call BASS_ChannelSetPosition(Mp3(2).chan, Mp3(2).file.SongStart)
    Call BASS_ChannelSlideAttributes(BASS_FX_TempoGetResampledHandle(Mp3(2).chan), -1, -2, -101, 1000)
    
'    Call BASS_FX_TempoStopAndFlush(Mp3(2).chan)
    Mp3(2).Estado = STOPED
    Label2.Caption = "DECK A : STOPED"
    analyzer.Refresh
    Timer1.Enabled = False
End Function

Public Function PauseMP3()
    Call BASS_ChannelPause(BASS_FX_TempoGetResampledHandle(Mp3(2).chan))
    Timer1.Enabled = False
    Mp3(2).Estado = PAUSED
    Label2.Caption = "DECK A : PAUSED"
End Function
Public Function UnloadMP3()
    'kung naga tokar hay e estap anay
    If Mp3(2).Estado = PLAYING Then StopMP3
    'kag e dis karga ang mp3
    BASS_StreamFree Mp3(2).chan
    Player_A(0).Caption = ""
    Player_A(1).Caption = ""
    Player_A(2).Caption = "00:00:00"
    Player_A(3).Caption = "00:00:00"
    Player_A(4).Caption = "00:00:00"
    Player_A(5).Caption = "00:00:00"
    lblBPM.Caption = ""
    lblAlbum.Caption = ""
    lblGenre.Caption = ""
    
    Mp3(2).Estado = IDLE
    Label2.Caption = "DECK A : IDLE"
    pos.value = 0
    LoadPoint = 0
    analyzer.Cls
End Function
Public Function LoadMP3()
Dim file As String
Dim length As Long
Dim idx As Integer
    'nothing to load then exit
    If Master.ListView2.ListItems.Count < 1 Then Exit Function
    If Timer1.Enabled = True Then
        Timer1.Enabled = False
    End If
    If IsDrag = False Then PlSourceRow = 1
    'load the file from the list
    idx = Master.ListView2.ListItems.Item(PlSourceRow).SubItems(4)
    Master.ListView2.ListItems.Remove (PlSourceRow)
    
    'save time and date last play
    MUSIC(idx).LastPlay = Now
    
    'save loaded song to history list
    Dim Lst As ListItem
    Set Lst = Master.ListView3.ListItems.Add()
       Lst.SmallIcon = 1
       Lst.text = MUSIC(idx).Title
       Lst.SubItems(1) = MUSIC(idx).Artist
       Lst.SubItems(2) = MUSIC(idx).bpm
       Lst.SubItems(3) = idx
    If Master.ListView2.ListItems.Count < 1 Then Master.ReloadHistory
    
    With Mp3(2)
        .file = MUSIC(idx)
        Player_A(0).Caption = .file.Title
        Player_A(1).Caption = .file.Artist
        lblAlbum.Caption = .file.Album
        lblGenre.Caption = .file.Genre
        .orgBPM = .file.bpm
    End With
    
    'buhian ang dati nga stream
    Call BASS_FX_DSP_Remove(Mp3(2).chan, BASS_FX_DSPFXVOLUME)
    Call BASS_FX_DSP_Remove(Mp3(2).chan, BASS_FX_DSPFX_PEAKEQ)
    Call BASS_FX_DSP_Remove(Mp3(2).chan, BASS_FX_DSPFX_FLANGER2)
    Call BASS_FX_BPM_Free(Mp3(2).chan)         'free the callback bpm
    Call BASS_FX_BPM_Free(Mp3(2).bpmhandle)    'free the decoding bpm
    Call BASS_FX_TempoFree(Mp3(2).chan)
    BASS_StreamFree Mp3(2).chan
    

    '----------------MP3-----------------------------
'    Mp3(2).chan = BASS_StreamCreateFile(BASSFALSE, Mp3(2).file.file, 0, 0, BASS_SAMPLE_LOOP Or BASS_STREAM_DECODE)
    Mp3(2).chan = BASS_StreamCreateFile(BASSFALSE, Mp3(2).file.file, 0, 0, BASS_STREAM_DECODE)
    
    Call BASS_ChannelGetAttributes(Mp3(2).chan, freq, vbNull, vbNull)
    Call BASS_FX_TempoCreate(Mp3(2).chan, 0)
      
      
      '--Retrieve Stream info---
    length = BASS_StreamGetLength(Mp3(2).chan)
    BASS_ChannelSetPosition Mp3(2).chan, length - 1
       'set pitch
    With pitchSlider
        .max = 1
        .Min = 0
        .max = (maxp - 1) * 100
        .Min = (minp - 1) * 100
        .value = 0
    End With
    
    'length of music
    Player_A(3).Caption = Atime.GetTime(Mp3(2).file.Len)
    'trigger point time
    Player_A(5).Caption = Atime.GetTime(BASS_ChannelBytes2Seconds(Mp3(2).chan, Mp3(2).file.SongEnd))
    'e preparar ang parametric ekwalayser
    Call EqualizerFrm.EqEnableB_Click
    'set the volume
    pos.max = BASS_StreamGetLength(Mp3(2).chan)
    pos.Min = 0
    pos.value = 0
    'get the load point
    LoadPoint = pos.max - (pos.max / 2)
    Mp3(2).Estado = READY
    Label2.Caption = "DECK A : READY"
    'StopMP3 'set the player ready
    Call Master.UpdateTimePlay
    'set the mixing point position
    SNYC2 = BASS_ChannelSetSync(Mp3(2).chan, BASS_SYNC_POS, Mp3(2).file.SongEnd, AddressOf modPublic.EndSync, 1)    ' set end sync
    SNYC2a = BASS_ChannelSetSync(Mp3(2).chan, BASS_SYNC_POS, Mp3(2).file.SongStart + 300000, AddressOf modPublic.UnloadSync, 1)   ' set end sync
    Mp3(1).FadeComplete = False

    lblBPM.Caption = Mp3(2).orgBPM
    MixOut = False
    Player2.MeReadyToStop = False
    
    
    'lblBPM.Caption = DecodeBPM2(True, 0, 30, Mp3(2).file.file)
    'Call BASS_FX_BPM_CallbackSet(Mp3(2).chan, AddressOf GetBPM_Callback2, 5, 0, BASS_FX_BPM_MULT2)
    Call BASS_ChannelSetPosition(Mp3(2).chan, Mp3(2).file.SongStart)
    Call BASS_StreamPlay(BASS_FX_TempoGetResampledHandle(Mp3(2).chan), 0, BASS_STREAM_AUTOFREE)
    Call BASS_ChannelPause(BASS_FX_TempoGetResampledHandle(Mp3(2).chan))
    
    efx5 False
    efx5 True
    fx5.BackColor = RGB(0, 255, 0)
End Function

Public Function PlayMP3()
'kung wala load endi mag play
If Mp3(2).Estado = IDLE Then Exit Function
Timer1.interval = 5
Timer1.Enabled = True

    Select Case Mp3(2).Estado
    Case READY
        Call BASS_ChannelResume(BASS_FX_TempoGetResampledHandle(Mp3(2).chan))
    Case STOPED
        Call BASS_ChannelSetPosition(Mp3(2).chan, Mp3(2).file.SongStart)
        Call BASS_StreamPlay(BASS_FX_TempoGetResampledHandle(Mp3(2).chan), 0, BASS_STREAM_AUTOFREE)
        Call BASS_ChannelPause(BASS_FX_TempoGetResampledHandle(Mp3(2).chan))
        Call BASS_ChannelResume(BASS_FX_TempoGetResampledHandle(Mp3(2).chan))
    Case PAUSED
        Call BASS_ChannelResume(BASS_FX_TempoGetResampledHandle(Mp3(2).chan))
    End Select
    
    MixerFrm.CrossFader_ValueChanged
Mp3(2).Estado = PLAYING
Label2.Caption = "DECK A : PLAYING"
'update pitch view
Call pitchSlider_ValueChanged
MeReadyToStop = False
End Function


Private Sub analyzer_DragDrop(Source As Control, X As Single, Y As Single)
If Source.Name = "ListView2" Then
    IsDrag = True
    LoadMP3
    IsDrag = False
    Set Master.ListView2.DropHighlight = Nothing
End If
End Sub







Private Sub cmdResetBPM_Click()
If Mp3(2).Estado <> IDLE Then
    pitchSlider.value = 0
End If
End Sub

Private Sub DeckLoad_Click()
LoadMP3
End Sub

Private Sub DeckPlay_Click()
PlayMP3
End Sub


Private Sub Form_Load()
Mp3(2).Estado = IDLE
Me.Height = 2490
'PaintBackground Me, BgSrc
End Sub

Public Sub fx1_Click()
Mp3(2).fx1 = Not Mp3(2).fx1
If Mp3(2).fx1 = True Then
    efx1 True
    fx1.BackColor = RGB(0, 255, 0)
Else
    efx1 False
    fx1.BackColor = &HC0C0C0
End If
End Sub

Private Sub fx2_Click()
Mp3(2).fx2 = Not Mp3(2).fx2
'///////////// Flanger //////////////
    If Mp3(2).fx2 = True Then
        efx2 True
        fx2.BackColor = RGB(0, 255, 0)
    Else
        efx2 False
        fx2.BackColor = &HC0C0C0
    End If
End Sub
Private Sub Picture5_DragDrop(Source As Control, X As Single, Y As Single)
If Source.Name = "ListView2" Then
    IsDrag = True
    LoadMP3
    IsDrag = False
    Set Master.ListView2.DropHighlight = Nothing
End If
End Sub

Private Sub fx3_Click()
Mp3(2).fx3 = Not Mp3(2).fx3
'///////////// brake //////////////
If Mp3(2).fx3 = True Then
        fx3.BackColor = RGB(0, 255, 0)
        efx3 True
Else
        efx3 False
        fx3.BackColor = &HC0C0C0
End If
End Sub

Private Sub fx4_Click()
Mp3(2).fx4 = Not Mp3(2).fx4
If Mp3(2).fx4 = True Then
     efx4 True
    fx4.BackColor = RGB(0, 255, 0)
Else
    efx4 False
    fx4.BackColor = &HC0C0C0
End If

End Sub

Private Sub fx5_Click()
Mp3(2).fx5 = Not Mp3(2).fx5
'///////////// brake //////////////
If Mp3(2).fx5 = True Then
        efx5 True
        fx5.BackColor = RGB(0, 255, 0)
Else
        efx5 False
        fx5.BackColor = &HC0C0C0
End If
End Sub





Private Sub pitchSlider_ValueChanged()
    Call BASS_FX_TempoSet(Mp3(2).chan, CLng(pitchSlider.value), -1, -100#)
    Mp3(2).newBPM = GetNewBPM2()
End Sub

Private Sub pos_MouseDown(Shift As Integer)
    Timer1.Enabled = False
End Sub

Private Sub pos_MouseUp(Shift As Integer)
    Call BASS_ChannelSetPosition(Mp3(2).chan, pos.value)
    Timer1.Enabled = True
End Sub

Private Sub StopDeck_Click()
StopMP3
End Sub
Public Sub Timer1_Timer()
If BASS_ChannelIsActive(Mp3(2).chan) = BASS_ACTIVE_PLAYING Then '
pos.value = BASS_ChannelGetPosition(Mp3(2).chan)
lblBPM.Caption = Mp3(2).newBPM

    '---------------- MP3 --------------------------
    If AutoDJ = True And Mp3(1).Estado = IDLE And pos.value >= LoadPoint Then
        Player1.LoadMP3
    End If
    
    If AutoDJ = True And MixOut = True And Player1.MeReadyToStop = True Then
        MixOut = False
        Player1.StopMP3
        Player1.UnloadMP3
        BASS_ChannelRemoveSync Mp3(2).chan, SNYC2a
    End If
    
    
    If AutoDJ = True And Mix = True Then
        Mix = False
        MeReadyToStop = True
        Player1.PlayMP3
        Call BASS_ChannelSlideAttributes(BASS_FX_TempoGetResampledHandle(Mp3(2).chan), -1, -2, -101, 1000)
        BASS_ChannelRemoveSync Mp3(2).chan, SNYC2
    End If



    If Master.enabSpectrum.Checked = True Then
    'spectrum analyzer BASS_DATA_FFT512
    BASS_ChannelGetData BASS_FX_TempoGetResampledHandle(Mp3(2).chan), SampleTemp(0), BASS_DATA_FFT512
    'display spectrum analyzer
    For i = 1 To 35
'        bitblt abuff.hdc, 4 * i, 65 - (Sqrt(SampleTemp(i)) * 140), 3, (Sqrt(SampleTemp(i)) * 140) + 1, gph.hdc, 0, 0, vbSrcCopy
        bitblt abuff.hdc, i * 5, 65 - (Sqrt(SampleTemp(i)) * 140), 4, (Sqrt(SampleTemp(i)) * 140) + 1, gph.hdc, 0, 65 - (Sqrt(SampleTemp(i)) * 140), vbSrcCopy
    
    Next i
    bitblt analyzer.hdc, 0, 0, 170, 100, abuff.hdc, 0, 0, vbSrcCopy
    abuff.Cls
    End If

Else
    Timer1.Enabled = False
End If
End Sub

Public Function Sqrt(ByVal num As Double) As Double
    Sqrt = num ^ 0.5
End Function

'effects
Private Sub efx1(state As Boolean)
Dim rv2 As BASS_FX_DSPREVERB
If state = True Then
    Call BASS_FX_DSP_Set(Mp3(2).chan, BASS_FX_DSPFX_REVERB, 1)
    Call BASS_FX_DSP_GetParameters(Mp3(2).chan, BASS_FX_DSPFX_REVERB, rv2)
        rv2.fLevel = 0.5
        rv2.lDelay = 3000
    Call BASS_FX_DSP_SetParameters(Mp3(2).chan, BASS_FX_DSPFX_REVERB, rv2)
Else
    Call BASS_FX_DSP_Remove(Mp3(2).chan, BASS_FX_DSPFX_REVERB)
End If
End Sub

Private Sub efx2(state As Boolean)
Dim fl2 As BASS_FX_DSPFLANGER2
    If state = True Then
        Call BASS_FX_DSP_Set(Mp3(2).chan, BASS_FX_DSPFX_FLANGER2, 1)
        Call BASS_FX_DSP_GetParameters(Mp3(2).chan, BASS_FX_DSPFX_FLANGER2, fl2)
        Call BASS_ChannelGetAttributes(Mp3(2).chan, fl2.lFreq, 0, 0)
               fl2.fDelay = 210 / 100
               fl2.fBPM = 120
               fl2.fWetDry = 2
        Call BASS_FX_DSP_SetParameters(Mp3(2).chan, BASS_FX_DSPFX_FLANGER2, fl2)
    Else
        Call BASS_FX_DSP_Remove(Mp3(2).chan, BASS_FX_DSPFX_FLANGER2)
    End If
End Sub

Private Sub efx3(state As Boolean)
Dim fl2 As BASS_FX_DSPECHO
    If state = True Then
        Call BASS_FX_DSP_Set(Mp3(2).chan, BASS_FX_DSPFX_ECHO, 1)
        Call BASS_FX_DSP_GetParameters(Mp3(2).chan, BASS_FX_DSPFX_ECHO, fl2)
            fl2.fLevel = 0.5
            fl2.lDelay = 8000
               
        Call BASS_FX_DSP_SetParameters(Mp3(2).chan, BASS_FX_DSPFX_ECHO, fl2)
    Else
        Call BASS_FX_DSP_Remove(Mp3(2).chan, BASS_FX_DSPFX_ECHO)
    End If
End Sub
Private Sub efx4(state As Boolean)
If state = True Then
    Call BASS_FX_TempoSet(Mp3(2).chan, -100#, -1, 8)
Else
    Call BASS_FX_TempoSet(Mp3(2).chan, -100#, -1, 0)
End If
End Sub

Private Sub efx5(state As Boolean)
Dim damp As BASS_FX_DSPDAMP
If state = True Then
        Call BASS_FX_DSP_Set(Mp3(2).chan, BASS_FX_DSPFX_DAMP, 1)
        Call BASS_FX_DSP_GetParameters(Mp3(2).chan, BASS_FX_DSPFX_DAMP, damp)
               damp.fGain = 1
               damp.fRate = 0.02
               damp.lDelay = 1000
               damp.lQuiet = 300
               damp.lTarget = 30000
        Call BASS_FX_DSP_SetParameters(Mp3(2).chan, BASS_FX_DSPFX_DAMP, damp)
Else
        Call BASS_FX_DSP_Remove(Mp3(2).chan, BASS_FX_DSPFX_DAMP)
End If
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
