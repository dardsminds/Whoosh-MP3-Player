VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEqualizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Private mvareGain As Single 'local copy
Private mvareBand1 As Single 'local copy
Private mvareBand2 As Single 'local copy
Private mvareBand3 As Single 'local copy
Private mvareBand4 As Single 'local copy
Private mvareBand5 As Single 'local copy
Private mvareBand6 As Single 'local copy
Private mvareBand7 As Single 'local copy
Private mvareBandWidth As Single 'local copy

Public Property Let eBandWidth(ByVal vData As Single)
    mvareBandWidth = vData
End Property
Public Property Get eBandWidth() As Single
    eBandWidth = mvareBandWidth
End Property


Public Sub EnableEqualizer(pHandler As Long, state As Boolean)
If state = True Then
    Dim f As Long
    Call BASS_ChannelGetAttributes(BASS_FX_TempoGetResampledHandle(pHandler), f, vbNull, vbNull)

    Call BASS_FX_DSP_Set(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, 0)
    Call BASS_FX_DSP_Set(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, 0)
    Call BASS_FX_DSP_Set(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, 0)
    Call BASS_FX_DSP_Set(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, 0)
    Call BASS_FX_DSP_Set(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, 0)
    Call BASS_FX_DSP_Set(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, 0)
    Call BASS_FX_DSP_Set(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, 0)
    Call BASS_FX_DSP_Set(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, 0)
    
    Dim EQ As BASS_FX_DSPPEAKEQ
    EQ.lFreq = f
    EQ.fBandwidth = Me.eBandWidth
    EQ.fGain = Me.egain

    EQ.lBand = 0
    EQ.fCenter = Me.eBand1
    Call BASS_FX_DSP_SetParameters(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, EQ)

    EQ.lBand = 1
    EQ.fCenter = Me.eBand2
    Call BASS_FX_DSP_SetParameters(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, EQ)

    EQ.lBand = 2
    EQ.fCenter = Me.eBand3
    Call BASS_FX_DSP_SetParameters(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, EQ)

    EQ.lBand = 3
    EQ.fCenter = Me.eBand4
    Call BASS_FX_DSP_SetParameters(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, EQ)

    EQ.lBand = 4
    EQ.fCenter = Me.eBand5
    Call BASS_FX_DSP_SetParameters(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, EQ)

    EQ.lBand = 5
    EQ.fCenter = Me.eBand6
    Call BASS_FX_DSP_SetParameters(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, EQ)

    EQ.lBand = 6
    EQ.fCenter = Me.eBand7
    Call BASS_FX_DSP_SetParameters(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, EQ)

Else
    Call BASS_FX_DSP_Remove(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ)
End If
End Sub

Public Sub Equalizer_Update(pHandler As Long, Index As Integer, gain As Integer)
    Dim EQ As BASS_FX_DSPPEAKEQ
    EQ.lBand = Index    'Band values you would like to get
    Call BASS_FX_DSP_GetParameters(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, EQ)
    EQ.fGain = Me.egain - gain
    Call BASS_FX_DSP_SetParameters(BASS_FX_TempoGetResampledHandle(pHandler), BASS_FX_DSPFX_PEAKEQ, EQ)
End Sub

Public Property Let eBand7(ByVal vData As Single)
    mvareBand7 = vData
End Property
Public Property Get eBand7() As Single
    eBand7 = mvareBand7
End Property
Public Property Let eBand6(ByVal vData As Single)
    mvareBand6 = vData
End Property
Public Property Get eBand6() As Single
    eBand6 = mvareBand6
End Property
Public Property Let eBand5(ByVal vData As Single)
    mvareBand5 = vData
End Property
Public Property Get eBand5() As Single
    eBand5 = mvareBand5
End Property
Public Property Let eBand4(ByVal vData As Single)
    mvareBand4 = vData
End Property
Public Property Get eBand4() As Single
    eBand4 = mvareBand4
End Property
Public Property Let eBand3(ByVal vData As Single)
    mvareBand3 = vData
End Property
Public Property Get eBand3() As Single
    eBand3 = mvareBand3
End Property
Public Property Let eBand2(ByVal vData As Single)
    mvareBand2 = vData
End Property
Public Property Get eBand2() As Single
    eBand2 = mvareBand2
End Property
Public Property Let eBand1(ByVal vData As Single)
    mvareBand1 = vData
End Property
Public Property Get eBand1() As Single
    eBand1 = mvareBand1
End Property
Public Property Let egain(ByVal vData As Single)
    mvareGain = vData
End Property
Public Property Get egain() As Single
    egain = mvareGain
End Property



