Attribute VB_Name = "modloop"
Public loop_(2) As Long     'loop start & end

Sub LoopSyncProc(ByVal handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
    BASS_ChannelSetPosition channel, loop_(0)
End Sub

