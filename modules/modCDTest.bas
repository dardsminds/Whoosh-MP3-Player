Attribute VB_Name = "modCDTest"
'////////////// CDDA module //////////////////////////////////////////////////////

Public Const MAXDRIVES = 10
Public curdrive As Long
Public stream(MAXDRIVES) As Long
Public seeking As Long

' End sync
Public Sub EndSync(ByVal Handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
'    If (frmCDTest.chkAdvance.value) Then  ' advance onto next track
'        Dim track As Long, drive As Long
'        track = BASS_CD_StreamGetTrack(channel)
'        drive = GetHiWord(track)
'        track = (GetLoWord(track) + 1) Mod BASS_CD_GetTracks(drive)
'        If (drive = curdrive) Then frmCDTest.lstTracks.ListIndex = track
'        Call PlayTrack(drive, track)
'    End If
End Sub

Public Sub PlayTrack(ByVal drive As Long, ByVal track As Long)
    On Error Resume Next    'to skip .sldPos.max error if stream(drive)=0
    stream(drive) = BASS_CD_StreamCreate(drive, track, BASS_CD_FREEOLD)  ' create stream
    'Call BASS_CD_ChannelSetSync(stream(drive), BASS_SYNC_END, 0, AddressOf EndSync, 0) ' set end sync
    Call BASS_StreamPlay(stream(drive), 0, 0) ' start playing
End Sub

Public Sub UpdateTrackList()
    Dim a As Long, tc As Long, L As Long, cdtext As Long, text As String
    Dim Lst As ListItem

    tc = BASS_CD_GetTracks(curdrive)

    If (tc = -1) Then Exit Sub  'no CD
    
    cdtext = BASS_CD_GetID(curdrive, BASS_CDID_TEXT) 'get CD-TEXT
    
    'add track on the list
    For a = 0 To tc - 1
        L = BASS_CD_GetTrackLength(curdrive, a)
        text = "Track " & Format(a + 1, "00")
      '  Master.PLayList.AddItem "" & Chr(9) & text & Chr(9) & "unknown" & Chr(9) & "unknown" & Chr(9) & "0" & Chr(9) & "0" & Chr(9) & "CD"
    'add item on the listview box
    Set Lst = Master.ListView2.ListItems.Add()
       Lst.SmallIcon = 1
       Lst.text = L
       Lst.SubItems(1) = text 'title
       Lst.SubItems(2) = "cdda"  'artist
       Lst.SubItems(3) = a    'bpm
       Lst.SubItems(4) = L    'lenght
       Lst.SubItems(5) = 0    'song start
       Lst.SubItems(6) = 0    'sond end
       Lst.SubItems(7) = "CD" 'file
    
    Next a
End Sub

