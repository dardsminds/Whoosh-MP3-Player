VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BrowserFrm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Whoosh MP3 Browser 1.20"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
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
   Icon            =   "BrowserFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNewCat 
      Height          =   345
      Left            =   4410
      TabIndex        =   30
      Top             =   4980
      Width           =   3435
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   315
      Left            =   7920
      TabIndex        =   29
      ToolTipText     =   "New Category"
      Top             =   4980
      Width           =   885
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   330
      Left            =   4410
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   4980
      Width           =   3435
   End
   Begin MSComctlLib.StatusBar stat 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   26
      Top             =   6345
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   529
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8476
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkOverwrite 
      Caption         =   "Overwrite song if exist in the Media library"
      Height          =   210
      Left            =   4395
      TabIndex        =   25
      Top             =   5685
      Width           =   3405
   End
   Begin VB.CheckBox chkAltSong 
      Height          =   210
      Left            =   1350
      TabIndex        =   24
      ToolTipText     =   "enable/disable alternative songname"
      Top             =   5925
      Width           =   225
   End
   Begin VB.TextBox txtAltsongname 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1620
      TabIndex        =   22
      ToolTipText     =   "Alternative songname if the music file does not have genre information"
      Top             =   5865
      Width           =   2430
   End
   Begin VB.ComboBox lblGenre 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "BrowserFrm.frx":0442
      Left            =   2805
      List            =   "BrowserFrm.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   16
      Text            =   "lblGenre"
      Top             =   5520
      Width           =   1245
   End
   Begin VB.TextBox lblTitle 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   705
      TabIndex        =   15
      Top             =   4725
      Width           =   3345
   End
   Begin VB.TextBox lblArtist 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   705
      TabIndex        =   14
      Top             =   4980
      Width           =   3345
   End
   Begin VB.TextBox lblYear 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   705
      TabIndex        =   13
      Top             =   5505
      Width           =   1455
   End
   Begin VB.TextBox lblAlbum 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   705
      TabIndex        =   12
      Top             =   5235
      Width           =   3345
   End
   Begin VB.CommandButton cmdAddToPlayList 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add to Library"
      Height          =   315
      Left            =   6135
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Add music to media library"
      Top             =   4350
      Width           =   1455
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Play"
      Height          =   315
      Left            =   4125
      TabIndex        =   9
      ToolTipText     =   "Preview song"
      Top             =   4350
      Width           =   960
   End
   Begin VB.CommandButton cmdPreviewStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5100
      TabIndex        =   8
      ToolTipText     =   "Preview stop"
      Top             =   4350
      Width           =   1020
   End
   Begin VB.CommandButton cmdDeleteFile 
      Caption         =   "Del"
      Height          =   315
      Left            =   7920
      TabIndex        =   7
      ToolTipText     =   "Delete the selected file"
      Top             =   4350
      Width           =   885
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   300
      Left            =   7920
      TabIndex        =   6
      Top             =   6015
      Width           =   870
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00E49667&
      ForeColor       =   &H00EEE7C4&
      Height          =   3870
      Left            =   4110
      Pattern         =   "*.wav"
      TabIndex        =   5
      Top             =   450
      Width           =   4725
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00E49667&
      ForeColor       =   &H00EEE7C4&
      Height          =   4170
      Left            =   60
      TabIndex        =   4
      Top             =   465
      Width           =   3990
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E1DFD9&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   105
      Width           =   3285
   End
   Begin VB.CheckBox chkAddAll 
      Caption         =   "Add all music file above to Media library"
      Height          =   210
      Left            =   4380
      TabIndex        =   1
      Top             =   5430
      Width           =   3180
   End
   Begin VB.PictureBox picBgSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1590
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   0
      Top             =   6555
      Width           =   720
   End
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   285
      Left            =   4395
      TabIndex        =   2
      Top             =   6015
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      Height          =   210
      Left            =   4410
      TabIndex        =   28
      Top             =   4740
      Width           =   675
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Alt. Songname"
      Height          =   300
      Left            =   135
      TabIndex        =   23
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   225
      TabIndex        =   21
      Top             =   4725
      Width           =   510
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Artist:"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   135
      TabIndex        =   20
      Top             =   4995
      Width           =   585
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   165
      TabIndex        =   19
      Top             =   5505
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Genre: "
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   2205
      TabIndex        =   18
      Top             =   5625
      Width           =   600
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Album:"
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   75
      TabIndex        =   17
      Top             =   5235
      Width           =   525
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   $"BrowserFrm.frx":0446
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   3450
      TabIndex        =   11
      Top             =   45
      Width           =   5325
   End
End
Attribute VB_Name = "BrowserFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim previewchan As Long  'channel for previewing music before add to list

Dim OUTR_VAL As Integer
Dim INTR_VAL As Integer

Dim rsMusic As ADODB.Recordset
Dim rsCheck As ADODB.Recordset
Dim rsMedia As ADODB.Recordset


Dim fHandle As Long 'handle of file to decode
Dim fhandlep As Long 'duplicate file but to play
Dim bpmhandle As Long
Dim buffers(0 To 8192) As Integer 'Decoding Channel Data
Dim samples(0 To 4096) As Integer ' Left hand channel
Dim PL As MUSICFILE

'    DatabaseFile = File1.FileName
'If Len(Dir1.Path) > 3 Then  'not a root directory
'    SourceDir = Dir1.Path & "\"
'Else
'    SourceDir = Dir1.Path
'End If
'
'    Source = SourceDir & DatabaseFile
'    Name Source As SourceDir & txtNewName.text & ".mdb"
'    File1.Refresh
'    BoxRename.visible = False



Private Sub DisplayFolderList()
    Dim sqltxt As String
    
    cmbCategory.Clear
    
    cmbCategory.AddItem "~Advertisement"
    cmbCategory.AddItem "~Jingles"
    cmbCategory.AddItem "~Program"
    cmbCategory.AddItem "~StationID"
    cmbCategory.AddItem "~VoiceOverBackground"
    
       
    Set rsMusic = OpenRS("SELECT * FROM folder ORDER BY folder")
    Do While Not rsMusic.EOF
            cmbCategory.AddItem rsMusic.Fields!folder
            rsMusic.MoveNext
    Loop
End Sub


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


Public Function GetStartMix(file As String) As Long
     Dim buf(0 To 5000) As Integer
     Dim count As Long
     Dim schan As Long
     Dim a, b As Integer
     
     schan = BASS_StreamCreateFile(False, file, 0, 0, BASS_STREAM_DECODE)
     Do While (BASS_ChannelIsActive(schan))
           b = BASS_ChannelGetData(schan, buf(0), 5000)
           For a = 0 To 3000
           If buf(a) >= INTR_VAL Then Exit Do
           Next a
           DoEvents
     Loop
     count = (BASS_ChannelGetPosition(schan))
     
     PL.Len = BASS_ChannelBytes2Seconds(schan, BASS_ChannelGetLength(schan))
     BASS_StreamFree schan
     GetStartMix = count
End Function
Public Function GetEndMix(file As String) As Long
     Dim buf(0 To 5000) As Integer
     Dim count As Long
     Dim echan As Long
     Dim a, b As Integer
     
    Call BASS_FX_ReverseFree(echan)
    echan = BASS_StreamCreateFile(BASSFALSE, file, 0, 0, BASS_STREAM_DECODE Or BASS_SAMPLE_LOOP Or BASS_MP3_SETPOS)
    Call BASS_FX_ReverseCreate(echan, 2#, BASSTRUE, 0)

     Do While (BASS_ChannelIsActive(BASS_FX_ReverseGetReversedHandle(echan)))
           b = BASS_ChannelGetData(BASS_FX_ReverseGetReversedHandle(echan), buf(0), 5000)
           For a = 5 To 500
           If buf(a) >= OUTR_VAL Then Exit Do
           Next a
           DoEvents
     Loop
     prog1.value = 70

     count = (BASS_ChannelGetPosition(BASS_FX_ReverseGetReversedHandle(echan)))
     count = BASS_ChannelGetLength(echan) - count
     
     'Call BASS_FX_PitchStopAndFlush(BASS_FX_ReverseGetReversedHandle(echan))
     Call BASS_FX_ReverseFree(echan)
     GetEndMix = count
     
End Function


Private Sub chkAddAll_Click()
    If chkAddAll.value = vbChecked Then
        MsgBox "This will add your music on the selected category, if the list of music has different category, those music will forcely fall to selected category", vbOKOnly, "WARNING!!!"
    End If
End Sub

Private Sub cmdClose_Click()
    'Set Cn = Nothing
    'End
    
    SaveINI "Browser", "FilePath", File1.Path
    SaveINI "Browser", "DirPath", Dir1.Path
    
    Unload Me
End Sub

Private Sub cmdDeleteFile_Click()
Dim flname As String
    
If File1.listcount = 0 Then
    MsgBox "No file to delete!! :-)", vbOKOnly, "Error!!!"
    Exit Sub
End If

If File1.ListIndex = -1 Then
    MsgBox "No selected file to delete!! :-)", vbOKOnly, "Error!!!"
    Exit Sub
End If

    If Right(File1.Path, 1) = "\" Then
        flname = File1.Path & File1.List(File1.ListIndex)
    Else
        flname = File1.Path & "\" & File1.List(File1.ListIndex)
    End If

'---confirm for file deletion----
If MsgBox("Delete file: " + flname + " anyway? ", vbOKCancel, "file Delete confirmation") = vbCancel Then Exit Sub
    On Error GoTo DelError
    Kill flname  'delete the file
    File1.Refresh
    
    Exit Sub
DelError:
    MsgBox "File deletion error, operation aborted.", vbOKOnly, "Error!!!"
End Sub


Private Sub cmdNew_Click()
    Dim sqltxt As String
    Select Case cmdNew.Caption
    
    Case Is = "New"
        cmdNew.Caption = "Save"
        txtNewCat.visible = True
        txtNewCat.text = ""
        txtNewCat.SetFocus
    Case Is = "Save"
        cmdNew.Caption = "New"
        txtNewCat.visible = False
        sqltxt = "INSERT INTO folder(folder) VALUES('" & Trim(txtNewCat.text) & "')"
        ExecuteQuery sqltxt
        DisplayFolderList
        cmbCategory.text = Trim(txtNewCat.text)
        MainFrm.DisplayFolderList
    End Select
End Sub

Private Sub File1_Click()
Dim Tag As Mp3IDTag
Dim Mp3file As String
    
lblTitle.text = ""
lblArtist.text = ""
lblYear.text = ""
lblGenre.text = ""
lblAlbum.text = ""
    
If Right(File1.Path, 1) = "\" Then
   Mp3file = File1.Path & File1.List(File1.ListIndex)
Else
   Mp3file = File1.Path & "\" & File1.List(File1.ListIndex)
End If
GetTag Mp3file, Tag
With Tag


If Trim(.Artist) = "" Then
    lblArtist.text = .Songname
Else
    lblArtist.text = .Artist
End If

If Trim(.Songname) = "" Then
    lblTitle.text = .Artist
Else
    lblTitle.text = .Songname
End If



lblYear.text = .Year
lblGenre.text = GenreText(Val(.Genre))
lblAlbum.text = .Album
End With
        
     stat.Panels(2).text = ""

End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim pfile As String

    'play
    If KeyCode = 32 Then
        cmdPreview_Click
    End If
  
   'stop
    If KeyCode = 27 Then
        cmdPreviewStop_Click
    End If

End Sub

Private Sub Form_Load()
Dim R As Integer
Dir1.Path = ReadINI("Browser", "DirPath", "c:\")
File1.Path = ReadINI("Browser", "FilePath", "c:\")
File1.Pattern = "*.mp3;*.wav;*.tfr;*.mpeg"
File1.Refresh

'    AUDIO.CueDevice = ReadINI("OUTPUT", "CueDevice", "1")
'    AUDIO.EnableCueDevice = ReadINI("OUTPUT", "CueDeviceEnabled", "0")
'    AUDIO.EnableOnAirDevice = ReadINI("OUTPUT", "DeviceEnabled", "1")
'    AUDIO.MemoryBuffer = ReadINI("OUTPUT", "MemoryBuffer", "1.2")
'    AUDIO.OnAirDevice = ReadINI("OUTPUT", "Device", "1")
'    AUDIO.OutputQuality = ReadINI("OUTPUT", "Quality", "44100")
'
    OUTR_VAL = ReadINI("MIXPOINT", "Outr_val", "5000")
    INTR_VAL = ReadINI("MIXPOINT", "Intr_val", "1000")
'
'   If (BASS_Init(AUDIO.OnAirDevice, AUDIO.OutputQuality, 0, Me.hwnd, 0) = 0) Then
'        MsgBox "Error: Couldn't Initialize Digital Output #1", vbCritical, "Digital output"
'        End
'   End If
    

    GENINFO.DatabaseName = ReadINI("DATABASE", "DatabasePath", "data.mdb")
    
    OpenMDB
    If Cn.state <> adStateOpen Then
        MsgBox "Database not found, click ok to search for database.", vbOKOnly, "Database Error"
        FrmDatabaseDir.Show 1
        End
    End If


LoadGenre     'load the genre list

lblGenre.Clear
For R = 0 To UBound(gMatrix)
    lblGenre.AddItem gMatrix(R)
Next R

    DisplayFolderList

'stat.Panels(1).Text = "(" & UBound(MUSIC) & ") song(s) in Media Library"
'PaintBackground Me, picBgSrc

txtNewCat.visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ov As String
    Dim iv As String
    
    ov = OUTR_VAL
    iv = INTR_VAL
    
    Call SaveINI("MIXPOINT", "Outr_val", ov)
    Call SaveINI("MIXPOINT", "Intr_val", iv)


    BASS_ChannelStop previewchan
    BASS_StreamFree previewchan
'    Call BASS_FX_Free
'    Call BASS_Free
End Sub

Private Sub File1_DblClick()
cmdPreview_Click
End Sub


Private Sub Dir1_Click()
On Error Resume Next
File1.Path = Dir1.List(Dir1.ListIndex)
File1.Refresh
End Sub

Private Sub Dir1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dir1_Click
End Sub

Private Sub Drive1_Change()
On Error GoTo er
Dir1.Path = Drive1.Drive
Dir1.Refresh
Exit Sub
er:
MsgBox "Drive Unavailable", vbOKOnly, "Error"
End Sub

Private Sub cmdPreview_Click()
If File1.ListIndex = -1 Then Exit Sub 'kung wala may gin pili
Dim pfile As String
    If Right(File1.Path, 1) = "\" Then
        pfile = File1.Path & File1.List(File1.ListIndex)
    Else
        pfile = File1.Path & "\" & File1.List(File1.ListIndex)
    End If
'kung naga tokar hay e stop
If BASS_ChannelIsActive(previewchan) = BASS_ACTIVE_PLAYING Then BASS_ChannelStop previewchan
    BASS_StreamFree previewchan
    previewchan = BASS_StreamCreateFile(BASSFALSE, pfile, 0, 0, 0)
    Call BASS_ChannelPlay(previewchan, BASSTRUE)
    
cmdPreviewStop.Enabled = True
End Sub

Private Sub cmdPreviewStop_Click()
    BASS_ChannelStop previewchan
    BASS_StreamFree previewchan
    cmdPreviewStop.Enabled = False
End Sub

Private Sub cmdAddToPlayList_Click()
Dim fcount As Integer
Dim count As Integer
    fcount = File1.listcount
    
    If Trim(cmbCategory.text) = "" Then
        MsgBox "Please select category first", vbOKOnly, "Category Error!"
        Exit Sub
    End If
    
If chkAddAll.value = vbChecked Then
    For count = 1 To fcount
    File1.ListIndex = count - 1
    AddtoMediaLibary count - 1
    Next count
Else
    AddtoMediaLibary File1.ListIndex
End If
End Sub

Private Sub AddtoMediaLibary(Index As Integer)
Dim flname As String
Dim Tag As Mp3IDTag
Dim songIndex As Integer
Dim sqltxt As String
Dim tmpstr As String

cmdPreview.Enabled = False
cmdDeleteFile.Enabled = False
chkAddAll.Enabled = False
chkOverwrite.Enabled = False
chkAltSong.Enabled = False
txtAltsongname.Enabled = False

    stat.Panels(2).text = "Opening file..."
    flname = File1.List(File1.ListIndex)
    If Right(File1.Path, 1) = "\" Then
        PL.file = File1.Path & File1.List(Index)
    Else
        PL.file = File1.Path & "\" & File1.List(Index)
    End If

    If InStr(PL.file, "'") Then
        stat.Panels(2).text = "The filename has single quote character!, please rename"
        GoTo JMPDONE
    End If

    If InStr(PL.Title, "'") Then
        tmpstr = Replace(PL.Title, "'", "`")
        PL.Title = tmpstr
    End If

    If InStr(PL.Artist, "'") Then
        tmpstr = Replace(PL.Artist, "'", "`")
        PL.Artist = tmpstr
    End If


    '----save info---
    PL.Artist = lblArtist.text
    PL.Genre = lblGenre.text
    PL.Album = lblAlbum.text
    PL.Year = lblYear.text
    
    If chkAltSong.value = vbChecked Then
        PL.Title = Trim(lblTitle.text) & "_" & txtAltsongname.text
    Else
        PL.Title = lblTitle.text
    End If
    
   
    'check if it is already on the database
    '-----Song already in database--
    Set rsCheck = OpenRS("SELECT * FROM music WHERE file='" & PL.file & "' AND Category='" & Trim(PL.Category) & "'")
    If rsCheck.RecordCount <> 0 And chkOverwrite.value = vbUnchecked Then
        stat.Panels(2).text = "Song already exist, operation aborted!"
        GoTo JMPDONE
    End If
    
    cmdAddToPlayList.Enabled = False
    Me.MousePointer = 11 ' change mouse pointer to wait
        
    DoEvents
    prog1.value = 40
    stat.Panels(2).text = " Detecting start mix point.."
    PL.SongStart = GetStartMix(PL.file)
    PL.vocalstart = GetStartMix(PL.file)
    
    prog1.value = 60
    
    
    DoEvents
    stat.Panels(2).text = " Detecting end mix point.."
    PL.SongEnd = GetEndMix(PL.file)
   
       
    DoEvents
    stat.Panels(2).text = " Detecting BPM..."
    prog1.value = 80
    bpmhandle = BASS_StreamCreateFile(BASSFALSE, PL.file, 0, 0, BASS_STREAM_DECODE)
    PL.bpm = BASS_FX_BPM_DecodeGet(bpmhandle, 0, 30, 0, BASS_FX_BPM_BKGRND Or BASS_FX_BPM_MULT2, 0)
    PL.Rating = 0
        
        'select category
        PL.Category = cmbCategory.text
        
        prog1.value = 85
        
        
        'add item on the listview box
        If chkOverwrite.value = vbChecked And songIndex <> -1 Then
            stat.Panels(2).text = "Overwriting music in Media Library.."
            'MUSIC(songIndex) = PL
        Else
            'Set rsMedia = OpenRS("SELECT * FROM music")
            stat.Panels(2).text = "Adding Music to Media Library.."
            sqltxt = "INSERT INTO music(file,title,artist,SongStart,SongEnd,vocalstart,bpm,Len,Category,LastPlay,MixType,Rating,Genre,Album,SongYear) "
            sqltxt = sqltxt & " VALUES('" & QuoteReplace(PL.file) & " ',"
            sqltxt = sqltxt & "'" & QuoteReplace(PL.Title) & " ',"
            sqltxt = sqltxt & "'" & QuoteReplace(PL.Artist) & " ',"
            sqltxt = sqltxt & PL.SongStart & ","
            sqltxt = sqltxt & PL.SongEnd & ","
            sqltxt = sqltxt & PL.vocalstart & ","
            sqltxt = sqltxt & PL.bpm & ","
            sqltxt = sqltxt & PL.Len & ","
            sqltxt = sqltxt & "'" & Trim(PL.Category) & " ',"
            sqltxt = sqltxt & "'" & PL.LastPlay & " ',"
            sqltxt = sqltxt & "'" & PL.MixType & " ',"
            sqltxt = sqltxt & PL.Rating & ","
            sqltxt = sqltxt & "'" & PL.Genre & " ',"
            sqltxt = sqltxt & "'" & QuoteReplace(PL.Album) & " ',"
            sqltxt = sqltxt & "'" & QuoteReplace(PL.Year) & " ')"
            ExecuteQuery sqltxt
        End If

          
        stat.Panels(2).text = "Updating Music Library..."
        prog1.value = 90
        cmdAddToPlayList.Enabled = True
        Me.MousePointer = 0
        prog1.value = 100
       
        stat.Panels(2).text = ""
        'stat.Panels(1).Text = "(" & UBound(MUSIC) & ") song(s) in Media Library"
        
JMPDONE:
        prog1.value = 0
        Me.MousePointer = 0
        cmdPreview.Enabled = True
        cmdDeleteFile.Enabled = True
        chkAddAll.Enabled = True
        chkOverwrite.Enabled = True
        chkAltSong.Enabled = True
        txtAltsongname.Enabled = True
        
        PL.file = ""
        PL.Title = ""
        PL.Artist = ""
        PL.bpm = 0
        PL.Category = ""
        PL.Genre = ""
        PL.Year = ""
End Sub


Private Sub txtNewCat_LostFocus()
        cmdNew.Caption = "New"
        txtNewCat.visible = False
End Sub
