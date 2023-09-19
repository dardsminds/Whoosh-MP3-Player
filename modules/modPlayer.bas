Attribute VB_Name = "modPlayer"


Sub SYNCMIX(ByVal Handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
    Select Case AutoDJ
        Case Is = False
        Exit Sub
    End Select

    Select Case user
    Case Is = MainFrm.Player1.pUser
        Player2MixNext = True
    Case Is = MainFrm.Player2.pUser
        Player1MixNext = True
    End Select
End Sub

Sub SYNCEND(ByVal Handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
'    Select Case AutoDJ
'        Case Is = False
'        Exit Sub
'    End Select
'
'    Select Case user
'    Case Is = MainFrm.Player1.pUser
'        MainFrm.Player1.StopPlay
'        MainFrm.Player1.UnloadMP3
'    Case Is = MainFrm.Player2.pUser
'        MainFrm.Player2.StopPlay
'        MainFrm.Player2.UnloadMP3
'    End Select
End Sub

