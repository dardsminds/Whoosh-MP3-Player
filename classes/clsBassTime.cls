VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBassTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Function GetTime(ByVal Seconds As Long) As String
    If Seconds <= 0 Then
        GetTime = "00:00:00"
        Exit Function
    End If
    Dim Hour As Single, Min As Single, Sec As Single
    Hour = Seconds / 60 / 60
    Sec = Seconds Mod 60
    Min = (Hour - Int(Hour)) * 60
    GetTime = Format(Int(Hour), "00") & ":" & Format(Int(Min), "00") & ":" & Format(Int(Sec), "00")
End Function

Public Function GetPlayingPos(ByVal Handle As Long) As Integer
    On Error GoTo err_GetPlayingPos
        GetPlayingPos = BASS_ChannelBytes2Seconds(Handle, BASS_ChannelGetPosition(Handle))
        'next shows how to get playing position without bass built function
        'GetPlayingPos = BASS_ChannelGetPosition(Handle) / GetBytesPerSecond(Handle)
    Exit Function
err_GetPlayingPos:
    GetPlayingPos = 0
End Function

Public Function GetDuration(ByVal Handle As Long) As Long
End Function

Public Function GetBytesPerSecond(ByVal Handle As Long) As Long
End Function

Public Function GetBitsPerSecond(ByVal Handle As Long, FileLength As Long) As Long
End Function

Public Function GetMode(ByVal Handle As Long) As String
End Function

Public Function GetBits(ByVal Handle As Long) As Byte
End Function

Public Function GetFrequency(ByVal Handle As Long) As Long
End Function

Public Function GetDXver() As Byte
End Function

