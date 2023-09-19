Attribute VB_Name = "BASS_FX"
'=============================================================================
' BASS_FX 2.3 - Copyright (c) 2002-2007 (: JOBnik! :) [Arthur Aminov, ISRAEL]
'                                                     [http://www.jobnik.org]
'
'     bugs/suggestions/questions -> e-mail: bass_fx@jobnik.org
'    ----------------------------------------------------------
'
' NOTE: This module will only work with BASS_FX version 2.3.0.3/4
'       Check http://www.un4seen.com for any later versions of BASS_FX.BAS
'
' * Requires BASS & BASS.BAS 2.3.0.3 - available @ www.un4seen.com
'=============================================================================

' Error codes returned by BASS_ErrorGetCode
Public Const BASS_FX_ERROR_NODECODE = 100    ' Not a decoding channel
Public Const BASS_FX_ERROR_BPMINUSE = 101    ' BPM/Beat detection is in use

' Tempo / Reverse / BPM / Beat flag
Public Const BASS_FX_FREESOURCE = &H10000    ' Free the source handle as well?

' Automatically reset DSPs, BPM/Beat Callbacks when channel's position is set (BASS_Set/GetConfig option)
Public Const BASS_FX_CONFIG_DSP_RESET = &H10500

'=============================================================================================
'   D S P (Digital Signal Processing)
'=============================================================================================

'  Multi-channel order of each channel is as follows:
'   3 channels       left-front, right-front, center.
'   4 channels       left-front, right-front, left-rear/side, right-rear/side.
'   6 channels (5.1) left-front, right-front, center, LFE, left-rear/side, right-rear/side.
'   8 channels (7.1) left-front, right-front, center, LFE, left-rear/side, right-rear/side, left-rear center, right-rear center.

' DSP channels flags
Public Const BASS_FX_DSP_CHANALL = -1        ' all channels at once (as by default)
Public Const BASS_FX_DSP_CHANNONE = 0        ' disable an effect for all channels
Public Const BASS_FX_DSP_CHAN1 = 1           ' left-front channel
Public Const BASS_FX_DSP_CHAN2 = 2           ' right-front channel
Public Const BASS_FX_DSP_CHAN3 = 4           ' see above info
Public Const BASS_FX_DSP_CHAN4 = 8           ' see above info
Public Const BASS_FX_DSP_CHAN5 = 16          ' see above info
Public Const BASS_FX_DSP_CHAN6 = 32          ' see above info
Public Const BASS_FX_DSP_CHAN7 = 64          ' see above info
Public Const BASS_FX_DSP_CHAN8 = 128         ' see above info

' if you have more than 8 channels, use BASS_FX_DSP_CHANNEL_N(n) below

' DSP effects
Public Enum DSPFX
    BASS_FX_DSPFX_SWAP                       ' Swap or Remap channels       / MC
    BASS_FX_DSPFX_ROTATE                     ' A channels volume ping-pong  / STEREO Only!
    BASS_FX_DSPFX_ECHO                       ' Echo                         / 2C max
    BASS_FX_DSPFX_FLANGER                    ' Flanger                      / MC
    BASS_FX_DSPFX_VOLUME                     ' Volume                       / MC
    BASS_FX_DSPFX_PEAKEQ                     ' Peaking Equalizer            / MC
    BASS_FX_DSPFX_REVERB                     ' Reverb                       / 2C max
    BASS_FX_DSPFX_LPF                        ' Low Pass Filter              / MC
    BASS_FX_DSPFX_S2M                        ' Stereo 2 Mono                / STEREO Only!
    BASS_FX_DSPFX_DAMP                       ' Dynamic Amplification        / MC
    BASS_FX_DSPFX_AUTOWAH                    ' Auto WAH                     / MC
    BASS_FX_DSPFX_ECHO2                      ' Echo 2                       / MC
    BASS_FX_DSPFX_PHASER                     ' Phaser                       / MC
    BASS_FX_DSPFX_ECHO3                      ' Echo 3                       / MC
    BASS_FX_DSPFX_CHORUS                     ' Chorus                       / MC
    BASS_FX_DSPFX_APF                        ' All Pass Filter              / MC
    BASS_FX_DSPFX_COMPRESSOR                 ' Compressor                   / MC
    BASS_FX_DSPFX_DISTORTION                 ' Distortion                   / MC
End Enum

' Swap/Remap channels
Public Type BASS_FX_DSPSWAP
    lChansOrder As Long                      ' A pointer to an array of channels order (BASS_FX_DSP_CHANxxx not in use, 1st channel index=0)
End Type

' Echo
Public Type BASS_FX_DSPECHO
    fLevel As Single                         ' [0....1....n] linear
    lDelay As Long                           ' [1200..30000]
End Type

' Flanger
Public Type BASS_FX_DSPFLANGER
    fWetDry As Single                        ' [0....1....n] linear
    fSpeed As Single                         ' [0......0.09]
    lChannel As Long                         ' BASS_FX_DSP_CHANxxx flag/s
End Type

' Volume
' Global volume is set to 1.0. Channels volume can't be greater than global volume.
' To set a new global volume, set lChannel = 0
Public Type BASS_FX_DSPVOLUME
    lChannel As Long                         ' BASS_FX_DSP_CHANxxx flag/s or 0 for global volume control
    fVolume As Single                        ' [0....1....n] linear
End Type

' Peaking Equalizer
Public Type BASS_FX_DSPPEAKEQ
    lBand As Long                            ' [0...............n] more bands means more memory & cpu usage
    lFreq As Long                            ' [1Hz...........nHz] current samplerate
    fBandwidth As Single                     ' [0.1......4......n] in octaves - Q is not in use (but BW has a priority over Q)
    fQ As Single                             ' [0.......1.......n] the EE kinda definition (linear) - Bandwidth is not in use
    fCenter As Single                        ' [1Hz...info.freq/3] in Hz
    fGain As Single                          ' [-15dB...0...+15dB] in dB
    lChannel As Long                         ' BASS_FX_DSP_CHANxxx flag/s
End Type

' Reverb
Public Type BASS_FX_DSPREVERB
    fLevel As Single                         ' [0....1....n] linear
    lDelay As Long                           ' [1200..10000]
End Type

' Low Pass Filter
Public Type BASS_FX_DSPLPF
    lFreq As Long                            ' [400Hz<=........nHz] current samplerate
    fResonance As Single                     ' [0.1.............10]
    fCutOffFreq As Single                    ' [1Hz....info.freq/2] cutoff frequency
    lChannel As Long                         ' BASS_FX_DSP_CHANxxx flag/s
End Type

' Dynamic Amplification
Public Type BASS_FX_DSPDAMP
    fTarget As Single                        ' target volume level                      [0<......1] linear
    fQuiet As Single                         ' quiet  volume level                      [0.......1] linear
    fRate As Single                          ' amp adjustment rate                      [0.......1] linear
    fGain As Single                          ' amplification level                      [0...1...n] linear
    fDelay As Single                         ' delay in seconds before increasing level [0.......n] linear
    lChannel As Long                         ' BASS_FX_DSP_CHANxxx flag/s
End Type

' Auto WAH
Public Type BASS_FX_DSPAUTOWAH
    fDryMix As Single                        ' dry (unaffected) signal mix              [-2......2]
    fWetMix As Single                        ' wet (affected) signal mix                [-2......2]
    fFeedback As Single                      ' feedback                                 [-1......1]
    fRate As Single                          ' rate of sweep in cycles per second       [0<....<10]
    fRange As Single                         ' sweep range in octaves                   [0<....<10]
    fFreq As Single                          ' base frequency of sweep Hz               [0<...1000]
    lChannel As Long                         ' BASS_FX_DSP_CHANxxx flag/s
End Type

' Echo 2
Public Type BASS_FX_DSPECHO2
    fDryMix As Single                        ' dry (unaffected) signal mix              [-2......2]
    fWetMix As Single                        ' wet (affected) signal mix                [-2......2]
    fFeedback As Single                      ' feedback                                 [-1......1]
    fDelay As Single                         ' delay sec                                [0<......6]
    lChannel As Long                         ' BASS_FX_DSP_CHANxxx flag/s
End Type

' Phaser
Public Type BASS_FX_DSPPHASER
    fDryMix As Single                        ' dry (unaffected) signal mix              [-2......2]
    fWetMix As Single                        ' wet (affected) signal mix                [-2......2]
    fFeedback As Single                      ' feedback                                 [-1......1]
    fRate As Single                          ' rate of sweep in cycles per second       [0<....<10]
    fRange As Single                         ' sweep range in octaves                   [0<....<10]
    fFreq As Single                          ' base frequency of sweep                  [0<...1000]
    lChannel As Long                         ' BASS_FX_DSP_CHANxxx flag/s
End Type

' Echo 3
Public Type BASS_FX_DSPECHO3
    fDryMix As Single                        ' dry (unaffected) signal mix              [-2......2]
    fWetMix As Single                        ' wet (affected) signal mix                [-2......2]
    fDelay As Single                         ' delay sec                                [0<......6]
    lChannel As Long                         ' BASS_FX_DSP_CHANxxx flag/s
End Type

' Chorus
Public Type BASS_FX_DSPCHORUS
    fDryMix As Single                        ' dry (unaffected) signal mix              [-2......2]
    fWetMix As Single                        ' wet (affected) signal mix                [-2......2]
    fFeedback As Single                      ' feedback                                 [-1......1]
    fMinSweep As Single                      ' minimal delay ms                         [0<..<6000]
    fMaxSweep As Single                      ' maximum delay ms                         [0<..<6000]
    fRate As Single                          ' rate ms/s                                [0<...1000]
    lChannel As Long                         ' BASS_FX_DSP_CHANxxx flag/s
End Type

' All Pass Filter
Public Type BASS_FX_DSPAPF
    fGain As Single                          ' reverberation time                       [-1=<..<=1]
    fDelay As Single                         ' delay sec                                [0<....<=6]
    lChannel As Long                         ' BASS_FX_DSP_CHANxxx flag/s
End Type

' Compressor
Public Type BASS_FX_DSPCOMPRESSOR
    fThreshold As Single                     ' compressor threshold                     [0<=...<=1]
    fAttacktime As Single                    ' attack time ms                           [0<.<=1000]
    fReleasetime As Single                   ' release time ms                          [0<.<=5000]
    lChannel As Long                         ' BASS_FX_DSP_CHANxxx flag/s
End Type

' Distortion
Public Type BASS_FX_DSPDISTORTION
    fDrive As Single                         ' distortion drive                         [0<=...<=5]
    fDryMix As Single                        ' dry (unaffected) signal mix              [-5<=..<=5]
    fWetMix As Single                        ' wet (affected) signal mix                [-5<=..<=5]
    fFeedback As Single                      ' feedback                                 [-1<=..<=1]
    fVolume As Single                        ' distortion volume                        [0=<...<=2]
    lChannel As Long                         ' BASS_FX_DSP_CHANxxx flag/s
End Type

Public Declare Function BASS_FX_DSP_Set Lib "bass_fx.dll" (ByVal Handle As Long, ByVal dsp_fx As DSPFX, ByVal priority As Long) As Long
'   Set any chosen DSP effect to any handle.
'   handle   : stream/music/wma/cd/any other supported add-on format
'   dsp_fx   : FX you wish to use
'   priority : The priority of the new DSP, which determines it's position in the DSP chain
'              DSPs with higher priority are called before those with lower.
'   RETURN   : TRUE if created (0=error!)

Public Declare Function BASS_FX_DSP_Remove Lib "bass_fx.dll" (ByVal Handle As Long, ByVal dsp_fx As DSPFX) As Long
'   Remove chosen DSP effect.
'   handle : stream/music/wma/cd/any other supported add-on format
'   dsp_fx : FX you wish to remove
'   RETURN : TRUE if removed (0=error!)

Public Declare Function BASS_FX_DSP_SetParameters Lib "bass_fx.dll" (ByVal Handle As Long, ByVal dsp_fx As DSPFX, ByRef par As Any) As Long
'   Set the parameters of a DSP effect.
'   handle : stream/music/wma/cd/any other supported add-on format
'   dsp_fx : FX you wish to set parameters to
'   par    : Pointer to the parameter structure
'   RETURN : TRUE if ok (0=error!)

Public Declare Function BASS_FX_DSP_GetParameters Lib "bass_fx.dll" (ByVal Handle As Long, ByVal dsp_fx As DSPFX, ByRef par As Any) As Long
'   Retrieve the parameters of a DSP effect.
'   handle : stream/music/wma/cd/any other supported add-on format
'   dsp_fx : FX you wish to get parameters from
'   par    : Pointer to the parameter structure
'   RETURN : TRUE if ok (0=error!)

Public Declare Function BASS_FX_DSP_Reset Lib "bass_fx.dll" (ByVal Handle As Long, ByVal dsp_fx As DSPFX) As Long
'   Call this function before changing position to avoid *clicks*
'   handle : stream/music/wma/cd/any other supported add-on format
'   dsp_fx : FX you wish to reset parameters of
'   RETURN : TRUE if ok (0=error!)

'=============================================================================================
'   TEMPO / PITCH SCALING / SAMPLERATE
'=============================================================================================

' NOTE: 1. Supported ONLY - mono / stereo - channels
'       2. Enable Tempo supported flags in BASS_FX_TempoCreate and the others to source handle.

Public Declare Function BASS_FX_TempoCreate Lib "bass_fx.dll" (ByVal chan As Long, ByVal flags As Long) As Long
'   Creates a resampling stream from a decoding channel.
'   chan     : a source handle returned by:
'                   BASS_StreamCreateFile     : flags = BASS_STREAM_DECODE ...
'                   BASS_MusicLoad            : flags = BASS_MUSIC_DECODE ...
'                   BASS_WMA_StreamCreateFile : flags = BASS_STREAM_DECODE ...
'                   BASS_CD_StreamCreate      : flags = BASS_STREAM_DECODE ...
'                   * Any other add-on handle using a decoding channel.
'   flags    : BASS_SAMPLE_SOFTWARE/LOOP/3D/FX or BASS_STREAM_DECODE/AUTOFREE or
'              BASS_SPEAKER_xxx or BASS_FX_FREESOURCE
'   RETURN   : the tempo stream handle (0=error!)

Public Declare Function BASS_FX_TempoGetSource Lib "bass_fx.dll" (ByVal chan As Long) As Long
'   Get the source channel handle.
'   chan   : tempo stream handle
'   RETURN : the source channel handle (0=error!)

Public Declare Function BASS_FX_TempoSet Lib "bass_fx.dll" (ByVal chan As Long, ByVal tempo As Single, ByVal samplerate As Single, ByVal pitch As Single) As Long
'   Set new values to tempo/rate/pitch to change its speed.
'   chan       : tempo stream (or source channel) handle
'   tempo      : in Percents  [-95%..0..+5000%]                 (-100 = leave current)
'   samplerate : in Hz, but calculates by the same % as tempo   (   0 = leave current)
'   pitch      : in Semitones [-60....0....+60]                 (-100 = leave current)
'   RETURN     : TRUE if ok (0=error!)

Public Declare Function BASS_FX_TempoGet Lib "bass_fx.dll" (ByVal chan As Long, ByRef tempo As Single, ByRef samplerate As Single, ByRef pitch As Single) As Long
'   Get tempo/rate/pitch values.
'   chan       : tempo stream (or source channel) handle
'   tempo      : current tempo          (0 = don't retrieve it)
'   samplerate : current samplerate     (0 = don't retrieve it)
'   pitch      : current pitch          (0 = don't retrieve it)
'   RETURN     : TRUE if ok (0=error!)

Public Declare Function BASS_FX_TempoGetRateRatio Lib "bass_fx.dll" (ByVal chan As Long) As Single
'   Get the ratio of the resulting rate and source rate (the resampling ratio).
'   chan   : tempo stream (or source channel) handle
'   RETURN : the ratio(0 = Error!)

' Tempo options. You can change all of them in real-time.
' option                                                   value
' ------                                                   -----
Public Const BASS_FX_TEMPO_OPTION_USE_AA_FILTER = 0      ' TRUE / FALSE
Public Const BASS_FX_TEMPO_OPTION_AA_FILTER_LENGTH = 1   ' 32 default (8 .. 128 taps)
Public Const BASS_FX_TEMPO_OPTION_USE_QUICKALGO = 2      ' TRUE / FALSE
Public Const BASS_FX_TEMPO_OPTION_SEQUENCE_MS = 3        ' 82 default
Public Const BASS_FX_TEMPO_OPTION_SEEKWINDOW_MS = 4      ' 14 default
Public Const BASS_FX_TEMPO_OPTION_OVERLAP_MS = 5         ' 12 default

Public Declare Function BASS_FX_TempoSetOption Lib "bass_fx.dll" (ByVal chan As Long, ByVal option_ As Long, ByVal value As Long) As Long
'   Set tempo options, one option each call.
'   chan    : tempo stream (or source channel) handle
'   option_ : BASS_FX_TEMPO_OPTION_xxx
'   value   : as written above.
'   RETURN  : TRUE if ok (0=error!)

Public Declare Function BASS_FX_TempoGetOption Lib "bass_fx.dll" (ByVal chan As Long, ByVal option_ As Long) As Long
'   Get tempo options, one option each call.
'   chan    : tempo stream (or source channel) handle
'   option_ : BASS_FX_TEMPO_OPTION_xxx
'   Return  : option value (-1=error!)

'=============================================================================================
'   R E V E R S E
'=============================================================================================

' NOTE: 1. MODs won't load without BASS_MUSIC_PRESCAN flag.
'       2. Enable Reverse supported flags in BASS_FX_ReverseCreate and the others to source handle.

Public Declare Function BASS_FX_ReverseCreate Lib "bass_fx.dll" (ByVal chan As Long, ByVal dec_block As Single, ByVal flags As Long) As Long
'   Creates a Reversed stream from a decoding channel.
'   chan      : a source handle returned by:
'                   BASS_StreamCreateFile     : flags = BASS_STREAM_DECODE ...
'                   BASS_MusicLoad            : flags = BASS_MUSIC_DECODE Or BASS_MUSIC_PRESCAN ...
'                   BASS_WMA_StreamCreateFile : flags = BASS_STREAM_DECODE ...
'                   BASS_CD_StreamCreate      : flags = BASS_STREAM_DECODE ...
'                   * Other stream add-on formats if created as decoded channel.
'               * For better MP3/2/1 Reverse playback use: BASS_STREAM_PRESCAN flag.
'   dec_block : decode in # of seconds blocks...
'               larger blocks = less seeking overhead but larger spikes.
'   flags     : BASS_SAMPLE_SOFTWARE/LOOP/3D/FX or BASS_STREAM_DECODE/AUTOFREE or
'               BASS_SPEAKER_xxx or BASS_FX_FREESOURCE
'   RETURN    : the reverse stream handle (0=error!)

Public Declare Function BASS_FX_ReverseGetSource Lib "bass_fx.dll" (ByVal chan As Long) As Long
'   Get the source channel handle.
'   chan   : reverse stream handle
'   RETURN : the source channel handle (0=error!)

' Playback directions
Public Const BASS_FX_RVS_REVERSE = 0
Public Const BASS_FX_RVS_FORWARD = 1

Public Declare Function BASS_FX_ReverseSetDirection Lib "bass_fx.dll" (ByVal chan As Long, ByVal direction As Long) As Long
'   Change playback direction.
'   chan      : reverse stream (or source channel) handle
'   direction : playback direction: BASS_FX_RVS_REVERSE or BASS_FX_RVS_FORWARD
'   RETURN    : TRUE if ok (0=error!)

Public Declare Function BASS_FX_ReverseGetDirection Lib "bass_fx.dll" (ByVal chan As Long) As Long
'   Get playback direction.
'   chan   : reverse stream (or source channel) handle
'   RETURN : playback direction(-1=error!)

'=============================================================================================
'   B P M (Beats Per Minute)
'=============================================================================================

' NOTE: Supported only mono or stereo channels.

' bpm flags
Public Const BASS_FX_BPM_BKGRND = 1   'If in use, then you can do other stuff while detection's in process (BPM/Beat)
Public Const BASS_FX_BPM_MULT2 = 2    'If in use, then will auto multiply bpm by 2 (if BPM < MinBPM*2)

'-----------
' Option - 1 - Get BPM from a decoded channel
'--------------------------------------------
Public Declare Function BASS_FX_BPM_DecodeGet Lib "bass_fx.dll" (ByVal chan As Long, ByVal StartSec As Single, ByVal EndSec As Single, ByVal minMaxBPM As Long, ByVal flags As Long, ByVal proc As Long) As Single
'   Get the original BPM of a decoding channel.
'   chan      : a handle returned by:
'                   BASS_StreamCreateFile     : flags = BASS_STREAM_DECODE ...
'                   BASS_MusicLoad            : flags = BASS_MUSIC_DECODE Or BASS_MUSIC_PRESCAN ...
'                   BASS_WMA_StreamCreateFile : flags = BASS_STREAM_DECODE ...
'                   BASS_CD_StreamCreate      : flags = BASS_STREAM_DECODE ...
'                   * Any other add-on handle using a decoding channel.
'   startSec  : start detecting position in seconds
'   endSec    : end detecting position in seconds
'   minMaxBPM : set min & max bpm, e.g: MAKELONG(LOWORD.HIWORD), LO=Min, HI=Max. 0 = defaults 45/230
'   flags     : BASS_FX_BPM_xxx or BASS_FX_FREESOURCE
'   proc      : user defined function to receive the process in percents, use 0 if not in use
'   RETURN    : the original BPM value (-1=error!)

'-----------
' Option - 2 - Auto get BPM by period of time in seconds
'-------------------------------------------------------
Public Declare Function BASS_FX_BPM_CallbackSet Lib "bass_fx.dll" (ByVal Handle As Long, ByVal proc As Long, ByVal period As Single, ByVal minMaxBPM As Long, ByVal flags As Long, ByVal user As Long) As Long
'   Enable getting BPM value by period of time in seconds.
'   handle    : stream/music/wma/cd/any other supported add-on format
'   proc      : user defined function to receive the bpm value
'   period    : detection period in seconds
'   minMaxBPM : set min & max bpm, e.g: MAKELONG(LOWORD.HIWORD), LO=Min, HI=Max. 0 = defaults 45/230
'   flags     : only BASS_FX_BPM_MULT2 flag is used
'   user      : user instance data to pass to the callback function.
'   RETURN    : TRUE if ok (0=error!)

Public Declare Function BASS_FX_BPM_CallbackReset Lib "bass_fx.dll" (ByVal Handle As Long) As Long
'   Reset the buffers. Call this function after changing position.
'   handle : stream/music/wma/cd/any other supported add-on format
'   RETURN : TRUE if ok (0=error!)

'---------------------------------------
'  Functions to use with Both options.
'---------------------------------------

' translation options
Public Const BASS_FX_BPM_X2 = 0         ' Multiply the original BPM value by 2 (may be called only once & will change the original BPM as well!)
Public Const BASS_FX_BPM_2FREQ = 1      ' BPM value to Frequency
Public Const BASS_FX_BPM_FREQ2 = 2      ' Frequency to BPM value
Public Const BASS_FX_BPM_2PERCENT = 3   ' BPM value to Percents
Public Const BASS_FX_BPM_PERCENT2 = 4   ' Percents to BPM value

Public Declare Function BASS_FX_BPM_Translate Lib "bass_fx.dll" (ByVal Handle As Long, ByVal val2tran As Single, ByVal trans As Long) As Single
'   Translate the given BPM to FREQ/PERCENT and vice versa or multiply BPM by 2.
'   handle   : stream/music/wma/cd/any other supported add-on format
'   val2tran : specify a value to translate to a given option (no matter if used X2).
'   trans    : any of the above translation option
'   RETURN   : new calculated value. (-1=error!)
'
'   NOTE     : This function will not detect the BPM, it will just translate the detected
'              original BPM value of a given handle.

Public Declare Sub BASS_FX_BPM_Free Lib "bass_fx.dll" (ByVal Handle As Long)
'   Free all resources used by a given handle (decode or callback bpm).
'   handle : stream/music/wma/cd/any other supported add-on format
'   RETURN : TRUE if ok (0=error!)
'
'   NOTE: If BASS_FX_FREESOURCE is used, then will free the source decoding channel as well.
'         You can't set/get this flag with BASS_ChannelSetFlags/BASS_ChannelGetInfo.

'=============================================================================================
'   B E A T
'=============================================================================================

' NOTE: Supported only mono or stereo channels.

'------------
' Real-time Beat position trigger functions
'-------------------------------------------
Public Declare Function BASS_FX_BPM_BeatCallbackSet Lib "bass_fx.dll" (ByVal Handle As Long, ByVal proc As Long, ByVal user As Long) As Long
'   Enable getting Beat position in seconds in real-time.
'   handle   : stream/music/wma/cd/any other supported add-on format
'   proc     : user defined BPMBEATPROC function to receive the beat position in seconds
'   user     : user instance data to pass to the callback function.
'   RETURN   : TRUE if ok (0=error!)

Public Declare Function BASS_FX_BPM_BeatCallbackReset Lib "bass_fx.dll" (ByVal Handle As Long) As Long
'   Reset the buffers. Call this function after changing position.
'   handle : stream/music/wma/cd/any other supported add-on format
'   RETURN : TRUE if ok (0=error!)

'------------
' Beat position detection functions
'-----------------------------------
Public Declare Function BASS_FX_BPM_BeatDecodeGet Lib "bass_fx.dll" (ByVal Handle As Long, ByVal StartSec As Single, ByVal EndSec As Single, ByVal flags As Long, ByVal proc As Long, ByVal user As Long) As Long
'   Enable getting Beat position in seconds of the decoded channel using the callback function.
'   chan     : stream/music/wma/cd/any other supported add-on format
'   startSec : start detecting position in seconds
'   endSec   : end detecting position in seconds
'   flags    : BASS_FX_BPM_BKGRND or BASS_FX_FREESOURCE
'   proc     : user defined function to receive the beat position in seconds
'   user     : user instance data to pass to the callback function.
'   RETURN   : TRUE if ok (0=error!)

'---------------------------------------
'  Functions to use with Both options.
'---------------------------------------
Public Declare Function BASS_FX_BPM_BeatSetParameters Lib "bass_fx.dll" (ByVal Handle As Long, ByVal bandwidth As Single, ByVal cutoffreq As Single, ByVal beat_rtime As Single) As Long
'   Set new values for beat detection.
'   handle     : stream/music/wma/cd/any other supported add-on format
'   bandwidth  : bandwidth in Hz [0<..<rate/2]      (-1.0 = leave current) [def. 10Hz]
'   cutoffreq  : cutoff frequency [0<..<rate/2]     (-1.0 = leave current) [def. 90Hz]
'   beat_rtime : beat release time in ms            (-1.0 = leave current) [def. 20ms]
'   RETURN     : TRUE if ok (0=error!)

Public Declare Function BASS_FX_BPM_BeatGetParameters Lib "bass_fx.dll" (ByVal Handle As Long, ByRef bandwidth As Single, ByRef cutoffreq As Single, ByRef beat_rtime As Single) As Long
'   Get current beat values.
'   handle     : stream/music/wma/cd/any other supported add-on format
'   bandwidth  : current bandwidth              (0 = don't retrieve it)
'   cutoffreq  : current cutoff frequency       (0 = don't retrieve it)
'   beat_rtime : current beat release time      (0 = don't retrieve it)
'   RETURN     : TRUE if ok (0=error!)          (0 = don't retrieve it)

Public Declare Sub BASS_FX_BPM_BeatFree Lib "bass_fx.dll" (ByVal Handle As Long)
'   Free all resources used by a given handle (decode or callback beat).
'   handle : stream/music/wma/cd/any other supported add-on format
'   RETURN : TRUE if ok (0=error!)
'
'   NOTE: If BASS_FX_FREESOURCE is used, then will free the source decoding channel as well.
'         You can't set/get this flag with BASS_ChannelSetFlags/BASS_ChannelGetInfo.


Public Sub BPMBEATPROC(ByVal Handle As Long, ByVal beatpos As Single, ByVal user As Long)
'   CALLBACK FUNCTION! - Uses with Beat Callback
'
'   Get the Beat position in seconds.
'   handle  : handle that the BASS_FX_BPM_BeatCallbackSet or BASS_FX_BPM_BeatDecodeGet has applied to
'   beatpos : the exact beat position in seconds
'   user    : the user instance data given when BASS_FX_BPM_BeatCallbackSet or BASS_FX_BPM_BeatDecodeGet was called
End Sub

Public Sub BPMPROCESSPROC(ByVal chan As Long, ByVal percent As Single)
'   CALLBACK FUNCTION! - Uses with BPM Option 1 (decode)
'
'   Get the detection process in percents of a channel.
'   chan    : channel that the BASS_FX_BPM_DecodeGet has applied to
'   percent : the process in percents [0%..100%]
End Sub

Public Sub BPMPROC(ByVal Handle As Long, ByVal bpm As Single, ByVal user As Long)
'   CALLBACK FUNCTION! - Uses with BPM Option 2 (callback)
'
'   Get the BPM after period of time in seconds.
'   handle : handle that the BASS_FX_BPM_CallbackSet has applied to
'   bpm    : the new original bpm value
'   user   : the user instance data given when BASS_FX_BPM_CallbackSet was called.
End Sub

' If you have more than 8 channels, use this macro
Public Function BASS_FX_DSP_CHANNEL_N(ByVal n As Long) As Long
    BASS_FX_DSP_CHANNEL_N = 2 ^ (n - 1)
End Function
