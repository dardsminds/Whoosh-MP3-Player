Attribute VB_Name = "modBPM"
Public Function GetNewBPM1() As Single
    'GetNewBPM1 = Format(Mp3(1).orgBPM * BASS_FX_TempoGetApproxPercents(Mp3(1).chan) / 100, "0.00")
End Function

Public Function GetNewBPM2() As Single
    'GetNewBPM2 = Format(Mp3(2).orgBPM * BASS_FX_TempoGetApproxPercents(Mp3(2).chan) / 100, "0.00")
End Function

Public Function DecodeBPM1(ByVal newStream As Boolean, ByVal StartSec As Single, ByVal EndSec As Single, file As String) As Single
'    If newStream Then
'        Mp3(1).bpmhandle = BASS_StreamCreateFile(BASSFALSE, file, 0, 0, BASS_STREAM_DECODE)
'    End If
'    Mp3(1).orgBPM = BASS_FX_BPM_DecodeGet(Mp3(1).bpmhandle, StartSec, EndSec, 0, BASS_FX_BPM_BKGRND Or BASS_FX_BPM_MULT2, 0)
'    Mp3(1).newBPM = GetNewBPM1()
'    DecodeBPM1 = Mp3(1).newBPM
End Function
Public Function DecodeBPM2(ByVal newStream As Boolean, ByVal StartSec As Single, ByVal EndSec As Single, file As String) As Single
'    If newStream Then
'        Mp3(2).bpmhandle = BASS_StreamCreateFile(BASSFALSE, file, 0, 0, BASS_STREAM_DECODE)
'    End If
'    Mp3(2).orgBPM = BASS_FX_BPM_DecodeGet(Mp3(2).bpmhandle, StartSec, EndSec, 0, BASS_FX_BPM_BKGRND Or BASS_FX_BPM_MULT2, 0)
'    Mp3(2).newBPM = GetNewBPM2()
'    DecodeBPM2 = Mp3(2).newBPM
End Function
'------------------------------------------
'----------- CALLBACK FUNCTIONS -----------
'------------------------------------------
'get the bpm after period of time
Public Sub GetBPM_Callback1(ByVal Handle As Long, ByVal bpm As Single)
'    Mp3(1).newBPM = GetNewBPM1()
End Sub
Public Sub GetBPM_Callback2(ByVal Handle As Long, ByVal bpm As Single)
'    Mp3(2).newBPM = GetNewBPM2()
End Sub

