VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsBassTime"
Attribute VB_Ext_KEY = "Member1" ,"clsEqualizer"
'local variable(s) to hold property value(s)
Private mvarfileName As String 'local copy
Private mvarpTitle As String 'local copy
Private mvarpArtist As String 'local copy
Private mvarpSongStart As Long 'local copy
Private mvarpSongEnd As Long 'local copy
Private mvarpBpm As Integer 'local copy
Private mvarpLen As Long 'local copy
Private mvarpCategory As String 'local copy
Private mvarpLastPlay As String 'local copy
Private mvarpMixType As String 'local copy
Private mvarpRating As Integer 'local copy
Private mvarpGenre As String 'local copy
Private mvarpAlbum As String 'local copy
Private mvarpYear As String 'local copy

Private phandle As Long
Private EndSync As Long
Private MixSync As Long

Private SampleBuff(0 To 300) As Single
Private endMixFired As Boolean
Private startMixFired As Boolean
Private LoadNextPointFired As Boolean

Private mvarclsEqualizer As clsEqualizer
Private mvarclsBassTime As clsBassTime

Private mvarisPlaying As Boolean 'local copy
Private loadnextpoint As Long
Private lsync As Long        'looping sync


Private mvarfxAGC As Boolean

Public Event onPlay()
Public Event onStop()
Public Event onPause()
Public Event onLoadMp3()
Public Event onUnloadMP3()
Public Event onEndMix()
Public Event onLoadNext()
Public Event onStartMix()

Private mvarIsLoaded As Boolean 'local copy
Private mvarpUser As Long 'local copy
Public Enum pStatus
    READY = 0
    STOPPED = 1
    PAUSED = 2
    PLAYING = 3
End Enum

Public Enum pMode
    pNORMAL = 0
    pBACKGROUND = 1
    pID = 2
    pPROGRAM = 3
End Enum

'local variable(s) to hold property value(s)
Public Status As pStatus
Public playerMode As pMode

Public Property Let Mode(ByVal vData As pMode)
    playerMode = vData
End Property

Public Property Get Mode() As pMode
    Mode = playerMode
End Property


Public Property Let pUser(ByVal vData As Long)
    mvarpUser = vData
End Property

Public Property Get pUser() As Long
    pUser = mvarpUser
End Property

Public Property Let IsLoaded(ByVal vData As Boolean)
    mvarIsLoaded = vData
End Property

Public Property Get IsLoaded() As Boolean
    IsLoaded = mvarIsLoaded
End Property

Public Sub EqualizerEnable(state As Boolean)
    Me.clsEqualizer.EnableEqualizer phandle, state
End Sub


Public Property Let fxAGC(ByVal vData As Boolean)
    mvarfxAGC = vData
    Dim damp As BASS_FX_DSPDAMP
    If mvarfxAGC = True Then
            Call BASS_FX_DSP_Set(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_DAMP, 1)
            Call BASS_FX_DSP_GetParameters(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_DAMP, damp)
                   damp.fGain = 1
                   damp.fRate = 0.02
                   damp.lDelay = 1000
                   damp.lQuiet = 300
                   damp.lTarget = 30000
            Call BASS_FX_DSP_SetParameters(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_DAMP, damp)
    Else
            Call BASS_FX_DSP_Remove(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_DAMP)
    End If
End Property
Public Property Get fxAGC() As Boolean
    fxAGC = mvarfxAGC
End Property

Public Function AudioLevel() As Long
    AudioLevel = BASS_ChannelGetLevel(BASS_FX_TempoGetResampledHandle(phandle))
End Function


Public Property Get clsBassTime() As clsBassTime
    Set clsBassTime = mvarclsBassTime
End Property

Public Property Set clsBassTime(vData As clsBassTime)
    Set mvarclsBassTime = vData
End Property

Public Function getMixpointTime() As String
    getMixpointTime = clsBassTime.GetTime(BASS_ChannelBytes2Seconds(phandle, Me.pSongEnd))
End Function

Public Function getLengthTime() As String
    getLengthTime = clsBassTime.GetTime(Me.pLen)
End Function

Public Function getCurrentTime() As String
    getCurrentTime = clsBassTime.GetTime(clsBassTime.GetPlayingPos(phandle))
End Function

Public Sub setTempo(ByVal tempo As Long)
'    Call BASS_FX_TempoSet(BASS_FX_TempoGet(phandle), CLng(tempo), -1, -100#)
'    Call BASS_FX_TempoSet(phandle, CLng(tempo), -1, -100#)
    'Mp3(1).newBPM = GetNewBPM1()
End Sub

Public Sub effectReverb(state As Boolean)
    Dim rv2 As BASS_FX_DSPREVERB
    If state = True Then
        Call BASS_FX_DSP_Set(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_REVERB, 1)
        Call BASS_FX_DSP_GetParameters(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_REVERB, rv2)
            rv2.fLevel = 0.5
            rv2.lDelay = 3000
        Call BASS_FX_DSP_SetParameters(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_REVERB, rv2)
    Else
        Call BASS_FX_DSP_Remove(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_REVERB)
    End If
End Sub

Public Sub effectFlanger(state As Boolean)
Dim fl2 As BASS_FX_DSPFLANGER2
    If state = True Then
        Call BASS_FX_DSP_Set(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_FLANGER2, 1)
        Call BASS_FX_DSP_GetParameters(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_FLANGER2, fl2)
        Call BASS_ChannelGetAttributes(BASS_FX_TempoGetResampledHandle(phandle), fl2.lFreq, 0, 0)
               fl2.fDelay = 210 / 100
               fl2.fBPM = 120
               fl2.fWetDry = 2
        Call BASS_FX_DSP_SetParameters(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_FLANGER2, fl2)
    Else
        Call BASS_FX_DSP_Remove(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_FLANGER2)
    End If
End Sub

Public Sub effectEcho(state As Boolean)
    Dim fl2 As BASS_FX_DSPECHO
    If state = True Then
        Call BASS_FX_DSP_Set(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_ECHO, 1)
        Call BASS_FX_DSP_GetParameters(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_ECHO, fl2)
            fl2.fLevel = 0.5
            fl2.lDelay = 8000
               
        Call BASS_FX_DSP_SetParameters(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_ECHO, fl2)
    Else
        Call BASS_FX_DSP_Remove(BASS_FX_TempoGetResampledHandle(phandle), BASS_FX_DSPFX_ECHO)
    End If
End Sub

Public Sub effectHighPitch(state As Boolean)
    If state = True Then
        Call BASS_FX_TempoSet(BASS_FX_TempoGetResampledHandle(phandle), -100#, -1, 8)
    Else
        Call BASS_FX_TempoSet(BASS_FX_TempoGetResampledHandle(phandle), -100#, -1, 0)
    End If
End Sub


Public Sub setPosition(ByVal pos As Long)
    Call BASS_ChannelSetPosition(phandle, pos)
End Sub

Public Function getStreamlenght() As Long
    getStreamlenght = BASS_ChannelGetLength(phandle)
End Function

Public Function getPosition() As Long
    If BASS_ChannelIsActive(phandle) = BASS_ACTIVE_PLAYING Then '
        getPosition = BASS_ChannelGetPosition(phandle)
        
        'If getPosition > Me.pSongEnd And endMixFired = False Then
        '    RaiseEvent onEndMix
        '    endMixFired = True
        'End If
        
        If AutoDJ = False Then Exit Function
        
        If getPosition > loadnextpoint And LoadNextPointFired = False Then
            RaiseEvent onLoadNext
            LoadNextPointFired = True
           
        End If
        
'        If getPosition > (loadnextpoint / 8) And startMixFired = False Then
        If getPosition > (Me.pSongStart + 100000) And startMixFired = False Then
            
            RaiseEvent onStartMix
            startMixFired = True
        End If
    Else
        getPosition = 0
    End If
End Function

Public Sub getSpectrum(outspec As PictureBox, specbuff As PictureBox, specsrc As PictureBox)
    BASS_ChannelGetData BASS_FX_TempoGetResampledHandle(phandle), SampleBuff(0), BASS_DATA_FFT512
    For i = 0 To 35
        BitBlt specbuff.hdc, i * 5, 65 - (Sqroot(SampleBuff(i)) * 140), 4, (Sqroot(SampleBuff(i)) * 140) + 1, specsrc.hdc, 0, 65 - (Sqroot(SampleBuff(i)) * 140), vbSrcCopy
    Next i
    BitBlt outspec.hdc, 0, 0, 170, 100, specbuff.hdc, 0, 0, vbSrcCopy
    specbuff.Cls

End Sub

Public Function GetRigthLevel() As Integer
    Dim level As Long
    level = BASS_ChannelGetLevel(BASS_FX_TempoGetResampledHandle(phandle))
    GetRigthLevel = HiWord(level)
End Function
Public Function GetLeftLevel() As Integer
    Dim level As Long
    level = BASS_ChannelGetLevel(BASS_FX_TempoGetResampledHandle(phandle))
    GetLeftLevel = LoWord(level)
End Function
Public Sub ClearSpectrum(outspec As PictureBox, specbuff As PictureBox, specsrc As PictureBox)
    outspec.Cls
    BitBlt outspec.hdc, 0, 0, 170, 100, specbuff.hdc, 0, 0, vbSrcCopy
    specbuff.Cls
End Sub

Private Function Sqroot(ByVal num As Double) As Double
    Sqroot = num ^ 0.5
End Function


Public Sub UnloadMP3()
Attribute UnloadMP3.VB_UserMemId = 0
    IsPlaying = False
    endMixFired = False
    RaiseEvent onUnloadMP3
    Me.IsLoaded = False
End Sub

Public Property Let IsPlaying(ByVal vData As Boolean)
    mvarisPlaying = vData
End Property

Public Property Get IsPlaying() As Boolean
    IsPlaying = mvarisPlaying
End Property

Private Sub Class_Initialize()
  Set mvarclsEqualizer = New clsEqualizer
  Set mvarclsBassTime = New clsBassTime
  
  clsEqualizer.eBand1 = 125
  clsEqualizer.eBand2 = 250
  clsEqualizer.eBand3 = 500
  clsEqualizer.eBand4 = 1000
  clsEqualizer.eBand5 = 2000
  clsEqualizer.eBand6 = 4000
  clsEqualizer.eBand7 = 8000
  
  clsEqualizer.eBandWidth = 1.5
  clsEqualizer.egain = 0
  
End Sub

Public Sub UpdateEqualizer(iband As Integer, egain As Integer)
    Me.clsEqualizer.Equalizer_Update phandle, iband, egain
End Sub

Public Property Get clsEqualizer() As clsEqualizer
    Set clsEqualizer = mvarclsEqualizer
End Property

Public Property Set clsEqualizer(vData As clsEqualizer)
    Set mvarclsEqualizer = vData
End Property
Private Sub Class_Terminate()
  Set mvarclsBassTime = Nothing
  Set mvarclsEqualizer = Nothing
End Sub

Public Sub setVolume(vol As Integer)
    Call BASS_ChannelSetAttributes(BASS_FX_TempoGetResampledHandle(phandle), -1, vol, -101)
End Sub

Public Function Load() As Boolean
    Call BASS_FX_DSP_Remove(phandle, BASS_FX_DSPFXVOLUME)
    Call BASS_FX_DSP_Remove(phandle, BASS_FX_DSPFX_PEAKEQ)
    Call BASS_FX_DSP_Remove(phandle, BASS_FX_DSPFX_FLANGER2)
    Call BASS_FX_BPM_Free(phandle)         'free the callback bpm
    Call BASS_FX_BPM_Free(phandle)    'free the decoding bpm
    Call BASS_FX_TempoFree(phandle)
    BASS_StreamFree phandle

    If Me.Mode = pBACKGROUND Then
        phandle = BASS_StreamCreateFile(BASSFALSE, mvarfileName, 0, 0, BASS_STREAM_DECODE Or BASS_SAMPLE_LOOP)
    Else
        phandle = BASS_StreamCreateFile(BASSFALSE, mvarfileName, 0, 0, BASS_STREAM_DECODE Or BASS_STREAM_AUTOFREE)
    End If

    Call BASS_ChannelGetAttributes(phandle, freq, vbNull, vbNull)
    Call BASS_FX_TempoCreate(phandle, 0)
    
    If EndSync <> 0 Then BASS_ChannelRemoveSync phandle, EndSync
    If MixSync <> 0 Then BASS_ChannelRemoveSync phandle, MixSync
    EndSync = BASS_ChannelSetSync(phandle, BASS_SYNC_END, Me.pSongEnd, AddressOf SYNCEND, Me.pUser)         ' set end sync
    MixSync = BASS_ChannelSetSync(phandle, BASS_SYNC_POS, Me.pSongEnd, AddressOf SYNCMIX, Me.pUser)        ' set end sync
    
    IsPlaying = False
    endMixFired = False
    LoadNextPointFired = False
    startMixFired = False
    loadnextpoint = Me.pSongEnd - (Me.pSongEnd / 2)
    Me.IsLoaded = True
    Me.Status = READY
    
    RaiseEvent onLoadMp3
End Function

Public Function Play() As Boolean
    If Me.Status = READY Then
        Call BASS_ChannelResume(BASS_FX_TempoGetResampledHandle(phandle))
    End If

    If Me.Status = PAUSED Then
        Call BASS_ChannelResume(BASS_FX_TempoGetResampledHandle(phandle))
    End If

    If Me.Status = STOPPED Then
        Select Case Me.Mode
        Case Is = pNORMAL
            Me.setPosition Me.pSongStart
            Call BASS_ChannelPlay(BASS_FX_TempoGetResampledHandle(phandle), 0, BASS_STREAM_AUTOFREE)
        Case Is = pBACKGROUND
            Me.setPosition Me.pSongStart
            Call BASS_ChannelPlay(BASS_FX_TempoGetResampledHandle(phandle), 0, BASS_SAMPLE_LOOP)
        Case Is = pPROGRAM
            Call BASS_ChannelPlay(BASS_FX_TempoGetResampledHandle(phandle), 0, BASS_STREAM_AUTOFREE)
        End Select
    End If
    
    'If mvarfxAGC = True Then  'BASS_SAMPLE_LOOP
        Me.fxAGC = False
        Me.fxAGC = True
    'End If
    
    
    IsPlaying = True
    endMixFired = False
    LoadNextPointFired = False
    startMixFired = False
    Me.Status = PLAYING
    RaiseEvent onPlay
End Function

Public Function ResumeMP3() As Boolean
    Call BASS_ChannelResume(BASS_FX_TempoGetResampledHandle(phandle))
    IsPlaying = True
    endMixFired = False
    LoadNextPointFired = False
    startMixFired = False
    Me.Status = PLAYING
    
    RaiseEvent onPlay
End Function

Public Function StopPlay()
   
    Select Case Me.Mode
    Case Is = pNORMAL
        Call BASS_ChannelSlideAttributes(BASS_FX_TempoGetResampledHandle(phandle), -1, -2, -101, 1500)
    Case Is = pBACKGROUND
        Call BASS_ChannelSlideAttributes(BASS_FX_TempoGetResampledHandle(phandle), -1, -2, -101, 500)
    Case Is = pPROGRAM
        Call BASS_ChannelSlideAttributes(BASS_FX_TempoGetResampledHandle(phandle), -1, -2, -101, 1500)
    End Select
    
    IsPlaying = False
    RaiseEvent onStop
    endMixFired = False
    LoadNextPointFired = False
    startMixFired = False
    Me.Status = STOPPED
End Function

Public Function PauseMp3()
    Call BASS_ChannelPause(BASS_FX_TempoGetResampledHandle(phandle))
    IsPlaying = False
    endMixFired = False
    LoadNextPointFired = False
    startMixFired = False
    Me.Status = PAUSED
    
    RaiseEvent onPause
End Function

Public Property Let pYear(ByVal vData As String)
    mvarpYear = vData
End Property
Public Property Get pYear() As String
    pYear = mvarpYear
End Property

Public Property Let pAlbum(ByVal vData As String)
    mvarpAlbum = vData
End Property

Public Property Get pAlbum() As String
    pAlbum = mvarpAlbum
End Property

Public Property Let pGenre(ByVal vData As String)
    mvarpGenre = vData
End Property

Public Property Get pGenre() As String
    pGenre = mvarpGenre
End Property


Public Property Let pRating(ByVal vData As Integer)
    mvarpRating = vData
End Property

Public Property Get pRating() As Integer
    pRating = mvarpRating
End Property

Public Property Let pMixType(ByVal vData As String)
    mvarpMixType = vData
End Property

Public Property Get pMixType() As String
    pMixType = mvarpMixType
End Property

Public Property Let pLastPlay(ByVal vData As String)
    mvarpLastPlay = vData
End Property

Public Property Get pLastPlay() As String
    pLastPlay = mvarpLastPlay
End Property

Public Property Let pCategory(ByVal vData As String)
    mvarpCategory = vData
End Property

Public Property Get pCategory() As String
    pCategory = mvarpCategory
End Property

Public Property Let pLen(ByVal vData As Long)
    mvarpLen = vData
End Property

Public Property Get pLen() As Long
    pLen = mvarpLen
End Property

Public Property Let pBpm(ByVal vData As Integer)
    mvarpBpm = vData
End Property

Public Property Get pBpm() As Integer
    pBpm = mvarpBpm
End Property

Public Property Let pSongEnd(ByVal vData As Long)
    mvarpSongEnd = vData
End Property

Public Property Get pSongEnd() As Long
    pSongEnd = mvarpSongEnd
End Property

Public Property Let pSongStart(ByVal vData As Long)
    mvarpSongStart = vData
End Property

Public Property Get pSongStart() As Long
    pSongStart = mvarpSongStart
End Property

Public Property Let pArtist(ByVal vData As String)
    mvarpArtist = vData
End Property

Public Property Get pArtist() As String
    pArtist = mvarpArtist
End Property

Public Property Let pTitle(ByVal vData As String)
    mvarpTitle = vData
End Property

Public Property Get pTitle() As String
    pTitle = mvarpTitle
End Property

Public Property Let pfile(ByVal vData As String)
    mvarfileName = vData
End Property

Public Property Get pfile() As String
    pfile = mvarfileName
End Property


Public Sub Playerloopeffect(stat As Boolean, startpos As Long, lenght As Long)
    If stat = True Then
        loop_(0) = startpos
        loop_(1) = loop_(0) + lenght
        Call BASS_ChannelRemoveSync(phandle, lsync) 'remove old sync
        BASS_ChannelSetPosition phandle, loop_(0)
        lsync = BASS_ChannelSetSync(phandle, BASS_SYNC_POS Or BASS_SYNC_MIXTIME, loop_(1), AddressOf LoopSyncProc, 0) 'set new sync
    Else
        Call BASS_ChannelRemoveSync(phandle, lsync) 'remove old sync
    End If
End Sub


Public Sub loopeffect1(stat As Boolean)
    If stat = True Then
        loop_(0) = Me.getPosition
        loop_(1) = loop_(0) + 20000
        Call BASS_ChannelRemoveSync(phandle, lsync) 'remove old sync
        BASS_ChannelSetPosition phandle, loop_(0)
        lsync = BASS_ChannelSetSync(phandle, BASS_SYNC_POS Or BASS_SYNC_MIXTIME, loop_(1), AddressOf LoopSyncProc, 0) 'set new sync
    Else
        Call BASS_ChannelRemoveSync(phandle, lsync) 'remove old sync
    End If
End Sub

Public Sub loopeffect2(stat As Boolean)
    If stat = True Then
        loop_(0) = Me.getPosition
        loop_(1) = loop_(0) + 40000
        Call BASS_ChannelRemoveSync(phandle, lsync) 'remove old sync
        BASS_ChannelSetPosition phandle, loop_(0)
        lsync = BASS_ChannelSetSync(phandle, BASS_SYNC_POS Or BASS_SYNC_MIXTIME, loop_(1), AddressOf LoopSyncProc, 0) 'set new sync
    Else
        Call BASS_ChannelRemoveSync(phandle, lsync) 'remove old sync
    End If
End Sub
Public Sub loopeffect3(stat As Boolean)
    If stat = True Then
        loop_(0) = Me.getPosition
        loop_(1) = loop_(0) + 80000
        Call BASS_ChannelRemoveSync(phandle, lsync) 'remove old sync
        BASS_ChannelSetPosition phandle, loop_(0)
        lsync = BASS_ChannelSetSync(phandle, BASS_SYNC_POS Or BASS_SYNC_MIXTIME, loop_(1), AddressOf LoopSyncProc, 0) 'set new sync
    Else
        Call BASS_ChannelRemoveSync(phandle, lsync) 'remove old sync
    End If
End Sub



