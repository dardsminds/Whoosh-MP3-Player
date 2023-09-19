VERSION 5.00
Begin VB.UserControl wplayer 
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   ScaleHeight     =   2040
   ScaleWidth      =   3975
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H008A8A8A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   0
      ScaleHeight     =   1740
      ScaleWidth      =   3750
      TabIndex        =   0
      Top             =   0
      Width           =   3750
      Begin VB.Label lblBPM 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2940
         TabIndex        =   18
         Top             =   870
         Width           =   765
      End
      Begin VB.Label lblRem 
         BackColor       =   &H00E49667&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2940
         TabIndex        =   17
         Top             =   660
         Width           =   765
      End
      Begin VB.Label lblCur 
         BackColor       =   &H00E49667&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2940
         TabIndex        =   16
         Top             =   450
         Width           =   765
      End
      Begin VB.Label lblTrig 
         BackColor       =   &H00E49667&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2940
         TabIndex        =   15
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lblLen 
         BackColor       =   &H00E49667&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2940
         TabIndex        =   14
         Top             =   30
         Width           =   765
      End
      Begin VB.Label lblGenre 
         BackColor       =   &H00E49667&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2670
         TabIndex        =   13
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Label lblAlbum 
         BackColor       =   &H00E49667&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   12
         Top             =   1500
         Width           =   1755
      End
      Begin VB.Label lblArtist 
         BackColor       =   &H00E49667&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   11
         Top             =   1290
         Width           =   3165
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00E49667&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   10
         Top             =   1080
         Width           =   3165
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
         TabIndex        =   9
         Top             =   885
         Width           =   195
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   480
         Width           =   195
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   1290
         Width           =   435
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
         TabIndex        =   2
         Top             =   1500
         Width           =   435
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
         TabIndex        =   1
         Top             =   1530
         Width           =   360
      End
   End
End
Attribute VB_Name = "wplayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type MSTRUCT
    file As String
    Title As String
    Artist As String
    SongStart As Long
    SongEnd As Long
    BPM As Integer
    Len As Long
    Category As String
    LastPlay As String
    MixType As String
    Rating As Integer
    Genre As String
    Album As String
    Year As String
End Type

Private WoPlayer As MSTRUCT
Private phandle As Long


Public Sub Load()
    lblTitle.Caption = WoPlayer.Title
    lblArtist.Caption = WoPlayer.Artist
    lblAlbum.Caption = WoPlayer.Album
    lblBPM.Caption = WoPlayer.BPM
    
    
    Call BASS_FX_DSP_Remove(phandle, BASS_FX_DSPFXVOLUME)
    Call BASS_FX_DSP_Remove(phandle, BASS_FX_DSPFX_PEAKEQ)
    Call BASS_FX_DSP_Remove(phandle, BASS_FX_DSPFX_FLANGER2)
    Call BASS_FX_BPM_Free(phandle)         'free the callback bpm
    Call BASS_FX_TempoFree(phandle)
    BASS_StreamFree phandle

    phandle = BASS_StreamCreateFile(BASSFALSE, WoPlayer.file, 0, 0, BASS_STREAM_DECODE Or BASS_STREAM_AUTOFREE)
    Call BASS_ChannelGetAttributes(phandle, freq, vbNull, vbNull)
    Call BASS_FX_TempoCreate(phandle, 0)
    
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "filename", WoPlayer.file, ""
    .WriteProperty "Title", WoPlayer.Title, ""
    .WriteProperty "Artist", WoPlayer.Artist, ""
    .WriteProperty "SongStart", WoPlayer.SongStart, "0"
    .WriteProperty "SongEnd", WoPlayer.SongEnd, "0"
    .WriteProperty "BPM", WoPlayer.BPM, "0"
    .WriteProperty "Category", WoPlayer.Category, "0"
    .WriteProperty "MixType", WoPlayer.MixType, "0"
    .WriteProperty "Rating", WoPlayer.Rating, "0"
End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    WoPlayer.file = .ReadProperty("filename", "")
    WoPlayer.Title = .ReadProperty("Title", "")
    WoPlayer.Artist = .ReadProperty("Artist", "")
    WoPlayer.SongStart = .ReadProperty("SongStart", "0")
    WoPlayer.SongEnd = .ReadProperty("SongEnd", "0")
    WoPlayer.BPM = .ReadProperty("BPM", "")
    WoPlayer.Category = .ReadProperty("Category", "")
    WoPlayer.MixType = .ReadProperty("MixType", "")
    WoPlayer.Rating = .ReadProperty("Rating", "0")
End With
End Sub

Public Property Get filename() As String
    filename = WoPlayer.file
End Property
Public Property Get Title() As String
    Title = WoPlayer.Title
End Property
Public Property Get Artist() As String
    Artist = WoPlayer.Artist
End Property
Public Property Get SongStart() As Long
    SongStart = WoPlayer.SongStart
End Property
Public Property Get SongEnd() As Long
    SongEnd = WoPlayer.SongEnd
End Property
Public Property Get BPM() As Single
    BPM = WoPlayer.BPM
End Property
Public Property Get Category() As String
    Category = WoPlayer.Category
End Property
Public Property Get MixType() As String
    MixType = WoPlayer.MixType
End Property
Public Property Get Rating() As Integer
    Rating = WoPlayer.Rating
End Property


Public Property Let filename(ByVal sNewValue As String)
    filename = sNewValue
    UserControl.PropertyChanged "filename"
End Property
Public Property Let Title(ByVal sNewValue As String)
    Title = sNewValue
    UserControl.PropertyChanged "Title"
End Property
Public Property Let Artist(ByVal sNewValue As String)
    Artist = sNewValue
    UserControl.PropertyChanged "Artist"
End Property
Public Property Let SongStart(ByVal sNewValue As Long)
    SongStart = sNewValue
    UserControl.PropertyChanged "SongStart"
End Property
Public Property Let SongEnd(ByVal sNewValue As Long)
    SongEnd = sNewValue
    UserControl.PropertyChanged "SongEnd"
End Property
Public Property Let BPM(ByVal sNewValue As Single)
    BPM = sNewValue
    UserControl.PropertyChanged "BPM"
End Property
Public Property Let Category(ByVal sNewValue As String)
    Category = sNewValue
    UserControl.PropertyChanged "Category"
End Property
Public Property Let MixType(ByVal sNewValue As String)
    MixType = sNewValue
    UserControl.PropertyChanged "MixType"
End Property
Public Property Let Rating(ByVal sNewValue As Integer)
    Rating = sNewValue
    UserControl.PropertyChanged "Rating"
End Property

