Attribute VB_Name = "Mp3Mod"
Option Explicit

Public Type MUSICFILE
    Index As Integer
    file As String
    Title As String
    Artist As String
    SongStart As Long
    SongEnd As Long
    vocalstart As Long
    bpm As Integer
    Len As Long
    Category As String
    LastPlay As String
    MixType As String
    Rating As Integer
    Genre As String
    Album As String
    Year As String
End Type

Public Type Mp3IDTag
      Songname As String * 30
      Artist As String * 30
      Album As String * 30
      Year As String * 4
      Comment As String * 30
      Genre As Integer
End Type


Private Type PROGRESS
    MaxLen As Single
    StartTrig As Single
    EndTrig As Single
    PosValue As Single
End Type

Public Type IRADIO
    StationName As String
    Address As String
    Category As String
End Type

Public MUSIC() As MUSICFILE
Public STATION() As IRADIO
Public gMatrix()



Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)


Public Function GetTag(FileName As String, ITag As Mp3IDTag) As Boolean
Dim f As Integer
Dim Tagg As String * 3
Dim temp As String
Dim i As Integer
Dim Cmd As String

f = FreeFile
On Error Resume Next
Open FileName For Binary As #f
Get #f, FileLen(FileName) - 127, Tagg

If Tagg = "TAG" Then
    GetTag = True
    Get #f, , ITag.Songname
    Get #f, , ITag.Artist
    Get #f, , ITag.Album
    Get #f, , ITag.Year
    Get #f, , ITag.Comment
    Get #f, , ITag.Genre
    
    ITag.Songname = Replace(ITag.Songname, Chr(0), " ", , , vbBinaryCompare)
    ITag.Artist = Replace(ITag.Artist, Chr(0), " ", , , vbBinaryCompare)
    ITag.Album = Replace(ITag.Album, Chr(0), " ", , , vbBinaryCompare)
    ITag.Year = Replace(ITag.Year, Chr(0), " ", , , vbBinaryCompare)
    ITag.Comment = Replace(ITag.Comment, Chr(0), " ", , , vbBinaryCompare)
    'ITag.Genre = ITag.Genre
    If Trim(ITag.Songname) = "" And Trim(ITag.Artist) = "" Then GoTo NOTAG
Else
NOTAG:
    GetTag = False
    temp = FileName
    If InStr(temp, "\") Then
        Do
        temp = Right(temp, Len(temp) - 1)
        Loop Until InStr(temp, "\") = 0
        
        If UCase(Right(temp, 4)) = ".MP3" Then temp = Left(temp, Len(temp) - 4)
        If InStr(temp, "-") Then
            ITag.Artist = Left(temp, InStr(temp, "-") - 1)
            ITag.Songname = Right(temp, Len(temp) - InStr(temp, "-") - 1)
        Else
            ITag.Songname = temp
        End If
    End If
End If
Close #f
End Function

Public Function FileExist(ByVal sPathName As String) As Boolean
    On Error Resume Next
    FileExist = (GetAttr(sPathName) And vbNormal) = vbNormal
    On Error GoTo 0
End Function

Public Function GenreText(Index As Integer) As String
    On Error GoTo Errhand
    GenreText = gMatrix(Index)
    Exit Function
Errhand:
    GenreText = ""
End Function

Sub AddMusic(newMusic As MUSICFILE)
    Dim i As Long
    ReDim Preserve MUSIC(UBound(MUSIC) + 1) As MUSICFILE
    MUSIC(UBound(MUSIC)) = newMusic
End Sub

Sub DeleteMusic(Index As Long)
    Dim i As Long
    For i = Index To UBound(MUSIC) - 1
        MUSIC(i) = MUSIC(i + 1) 'move data up higher
    Next
    'resize the buffer but preserve its contents
    ReDim Preserve MUSIC(UBound(MUSIC) - 1) As MUSICFILE
End Sub

Public Sub SetMusicCategory(Index As Long, newCat As String)
    MUSIC(Index).Category = newCat
End Sub

Public Sub SetMusicRating(Index As Long, newRating As Integer)
    MUSIC(Index).Rating = newRating
End Sub

Public Sub ReadPlayList()
    Dim bSize As Long
    Open App.Path & "\whoosh.pl" For Binary As #2
    Get #2, 1, bSize          ' Read the number of items.
    ReDim MUSIC(0 To bSize) As MUSICFILE
    Get #2, , MUSIC()
    Close #2
End Sub

Public Sub SavePlayList()
    Open App.Path & "\whoosh.pl" For Binary As #1
    Put #1, 1, CLng(UBound(MUSIC))
    Put #1, , MUSIC()
    Close #1
End Sub

Public Function FindMusic(Title As String, Artist As String) As Long
    Dim idx As Integer
    Dim found As Boolean
    found = False
    For idx = 0 To UBound(MUSIC)
    If MUSIC(idx).Title = Title And MUSIC(idx).Artist = Artist Then
        FindMusic = idx
        found = True
        Exit For
    End If
    Next idx
    If found = False Then
        FindMusic = -1
    End If
End Function

Public Sub SaveGenre()
'Dim gMatrix()
'gMatrix = Array("Blues", "Classic Rock", "Country", "Dance", "Disco", "Funk", "Grunge", _
'"Hip -Hop", "Jazz", "Metal", "New Age", "Oldies", "Other", "Pop", "RnB", "Rap", "Reggae", _
'"Rock", "Techno", "Industrial", "Alternative", "Ska", "Death Metal", "Pranks", _
'"Soundtrack", "Euro -Techno", "Ambient", "Trip -Hop", "Vocal", "Jazz Funk", "Fusion", _
'"Trance", "Classical", "Instrumental", "Acid", "House", "Game", "Sound Clip", "Gospel", _
'"Noise", "AlternRock", "Bass", "Soul", "Punk", "Space", "Meditative", "Instrumental Pop", _
'"Instrumental Rock", "Ethnic", "Gothic", "Darkwave", "Techno -Industrial", "Electronic", _
'"Pop -Folk", "Eurodance", "Dream", "Southern Rock", "Comedy", "Cult", "Gangsta", "Top 40", _
'"Christian Rap", "Pop/Funk", "Jungle", "Native American", "Cabaret", "New Wave", _
'"Psychadelic", "Rave", "Showtunes", "Trailer", "Lo -Fi", "Tribal", "Acid Punk", "Acid Jazz", _
'"Polka", "Retro", "Musical", "Rock & Roll", "Hard Rock", "Folk", "Folk/Rock", "National Folk", _
'"Swing", "Bebob", "Latin", "Revival", "Celtic", "Bluegrass", "Avantgarde", "Gothic Rock", _
'"Progressive Rock", "Psychedelic Rock", "Symphonic Rock", "Slow Rock", "Big Band", "Chorus", "Easy Listening", _
'"Acoustic", "Humour", "Speech", "Chanson", "Opera", "Chamber Music", "Sonata", "Symphony", "Booty Bass", _
'"Primus", "Porn Groove", "Satire", "Slow Jam", "Club", "Tango", "Samba", "Folklore", "Ballad", "Power Ballad", _
'"Rhythmic Soul", "Freestyle", "Duet", "Punk Rock", "Drum Solo", "A Cappella", _
'"Euro - House", "Dance Hall", "Goa", "Drum & Bass", "Club - House", "Hardcore", "Terror", "Indie", "BritPop", _
'"Negerpunk", "Polsk Punk", "Beat", "Christian Gangsta Rap", "Heavy Metal", "Black Metal", "Crossover", _
'"Contemporary Christian", "Christian Rock", "Merengue", "Salsa", "Thrash Metal", "Anime", "JPop", "Synthpop")
'Open App.Path & "\genre.dat" For Binary As #1
'Put #1, 1, CLng(UBound(gMatrix))
'Put #1, , gMatrix()
'Close #1
End Sub

Public Sub LoadGenre()
    Dim bSize As Long
    Open App.Path & "\genre.dat" For Binary As #2
    Get #2, 1, bSize          ' Read the number of items.
    ReDim gMatrix(0 To bSize)
    Get #2, , gMatrix()
    Close #2
End Sub
