Attribute VB_Name = "modEqualizer"
'Public Equalizer(7) As Long
'Public EqualizerB(7) As Long

'Public Sub EqualizerSet()
'If EqualizerFrm.EqEnable = 1 Then
'    'setup the EQ effects
'    Dim P As BASS_FXPARAMEQ
'    Dim i As Integer
'
'    Equalizer(0) = BASS_ChannelSetFX(Mp3(1).chan, BASS_FX_PARAMEQ)
'    Equalizer(1) = BASS_ChannelSetFX(Mp3(1).chan, BASS_FX_PARAMEQ)
'    Equalizer(2) = BASS_ChannelSetFX(Mp3(1).chan, BASS_FX_PARAMEQ)
'    Equalizer(3) = BASS_ChannelSetFX(Mp3(1).chan, BASS_FX_PARAMEQ)
'    Equalizer(4) = BASS_ChannelSetFX(Mp3(1).chan, BASS_FX_PARAMEQ)
'    Equalizer(5) = BASS_ChannelSetFX(Mp3(1).chan, BASS_FX_PARAMEQ)
'    Equalizer(6) = BASS_ChannelSetFX(Mp3(1).chan, BASS_FX_PARAMEQ)
'
'    P.fGain = 0
'    P.fBandwidth = Val(EqualizerFrm.txtBandWidth.text)
    
    'Set Equalizer Center Frequency
'    P.fCenter = Val(EqualizerFrm.EqCenterFreq(0).text) '1
'    Call BASS_FXSetParameters(Equalizer(0), P)
'    P.fCenter = Val(EqualizerFrm.EqCenterFreq(1).text) '2
''    Call BASS_FXSetParameters(Equalizer(1), P)
 '   P.fCenter = Val(EqualizerFrm.EqCenterFreq(2).text) '3
'    Call BASS_FXSetParameters(Equalizer(2), P)
'    P.fCenter = Val(EqualizerFrm.EqCenterFreq(3).text) '4
'    Call BASS_FXSetParameters(Equalizer(3), P)
'    P.fCenter = Val(EqualizerFrm.EqCenterFreq(4).text) '5
'    Call BASS_FXSetParameters(Equalizer(4), P)
'    P.fCenter = Val(EqualizerFrm.EqCenterFreq(5).text) '6
'    Call BASS_FXSetParameters(Equalizer(5), P)
'    P.fCenter = Val(EqualizerFrm.EqCenterFreq(6).text) '7
'    Call BASS_FXSetParameters(Equalizer(6), P)
    
    'update dx8 fx
'    For i = 0 To 6
'        UpdateEqualizer i
'    Next i
'Else
'    Call BASS_FX_DSP_Remove(Mp3(1).chan, BASS_FX_DSPFXEQ)
'End If
'End Sub

' Update DX8 PARAMETRIC EQ
'Public Sub UpdateEqualizer(ByVal b As Integer)
'    Dim P As BASS_FX_DSPEQ
'     Call BASS_FXGetParameters(Equalizer(b), P)
'    P.fGain = EqualizerFrm.EqGain.Value - EqualizerFrm.Equalizer(b).Value * -1
'    Call BASS_FXSetParameters(Equalizer(b), P)
'    Dim eq As BASS_FX_DSPEQ
    
'    P.eqBand = Index   'Band values you would like to get
    
'    Call BASS_FX_DSP_GetParameters(Mp3(1).chan, BASS_FX_DSPFXEQ, P)
'    P.EqGain = EqualizerFrm.EqGain.Value - EqualizerFrm.Equalizer(b).Value * -1
'    Call BASS_FX_DSP_SetParameters(Mp3(1).chan, BASS_FX_DSPFXEQ, P)
'End Sub

'Public Sub EqualizerSetB(ByVal Handle As Long)
    'setup the EQ effects
'    Dim P As BASS_FXPARAMEQ
'    Dim i As Integer
    
'    EqualizerB(0) = BASS_ChannelSetFX(Handle, BASS_FX_PARAMEQ)
'    EqualizerB(1) = BASS_ChannelSetFX(Handle, BASS_FX_PARAMEQ)
'    EqualizerB(2) = BASS_ChannelSetFX(Handle, BASS_FX_PARAMEQ)
'    EqualizerB(3) = BASS_ChannelSetFX(Handle, BASS_FX_PARAMEQ)
'    EqualizerB(4) = BASS_ChannelSetFX(Handle, BASS_FX_PARAMEQ)
'    EqualizerB(5) = BASS_ChannelSetFX(Handle, BASS_FX_PARAMEQ)
'    EqualizerB(6) = BASS_ChannelSetFX(Handle, BASS_FX_PARAMEQ)
      
'    P.fGain = 0
'    P.fBandwidth = Val(EqualizerFrm.txtBandWidth.text)
    
    'Set Equalizer Center Frequency
'    P.fCenter = Val(EqualizerFrm.EqCenterFreq(0).text) '1
'    Call BASS_FXSetParameters(EqualizerB(0), P)
'    P.fCenter = Val(EqualizerFrm.EqCenterFreq(1).text) '2
'    Call BASS_FXSetParameters(EqualizerB(1), P)
'    P.fCenter = Val(EqualizerFrm.EqCenterFreq(2).text) '3
'    Call BASS_FXSetParameters(EqualizerB(2), P)
'    P.fCenter = Val(EqualizerFrm.EqCenterFreq(3).text) '4
'    Call BASS_FXSetParameters(EqualizerB(3), P)
'    P.fCenter = Val(EqualizerFrm.EqCenterFreq(4).text) '5
'    Call BASS_FXSetParameters(EqualizerB(4), P)
'    P.fCenter = Val(EqualizerFrm.EqCenterFreq(5).text) '6
'    Call BASS_FXSetParameters(EqualizerB(5), P)
'    P.fCenter = Val(EqualizerFrm.EqCenterFreq(6).text) '7
'    Call BASS_FXSetParameters(EqualizerB(6), P)
    'update dx8 fx
'    For i = 0 To 6
'        UpdateEqualizerB i
'    Next i
'End Sub

' Update DX8 PARAMETRIC EQ
'Public Sub UpdateEqualizerB(ByVal b As Integer)
'    Dim P As BASS_FXPARAMEQ
'     Call BASS_FXGetParameters(EqualizerB(b), P)
'    P.fGain = EqualizerFrm.EqGain.Value - EqualizerFrm.EqualizerB(b).Value * -1
'    Call BASS_FXSetParameters(EqualizerB(b), P)
'End Sub

