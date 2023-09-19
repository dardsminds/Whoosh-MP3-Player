Attribute VB_Name = "main_browser"
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SetWindowPos Lib "user32" (ByVal h%, ByVal hb%, ByVal X%, ByVal Y%, ByVal CX%, ByVal CY%, ByVal f%) As Integer
'API calls used for reading and writing of preferences
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


' Win32 Declarations for DisableX
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&


Type BASSENGINE
    OnAirDevice As Integer
    CueDevice As Integer
    OutputQuality As Long
    MemoryBuffer As Single
    EnableOnAirDevice As Integer
    EnableCueDevice As Integer
End Type

'Public AutoDJ As Boolean
'Public Mix As Boolean
'Public MixOut As Boolean
Public PlSourceRow As Integer
Public IsDrag As Boolean
Public COMPRESSOR As BASS_FXCOMPRESSOR
Public MXP As MUSICFILE
Public AUDIO As BASSENGINE

'player timer
'Public Atime As CBASS_TIME
'Public Btime As CBASS_TIME

'Public Sample1 As Long
'Public Sample2 As Long
'Public Sample3 As Long
'Public Sample4 As Long

'Public GainAdj(0 To 200) As Integer
'Public knee As Integer
'Public release As Integer
'Public interval As Integer
Public CurrentCategory As String
Public GENINFO As DATABASE_INFO
Const INI_FILE = "Whoosh.ini"



' ---------------Save and Write to INI file function-----
Public Function SaveINI(Key As String, Section As String, sVal As String) As Long
  SaveINI = WritePrivateProfileString(Key, Section, sVal, App.Path & "\" & INI_FILE)
End Function

Public Function ReadINI(Key As String, Section As String, DefVal As String) As String
  Dim lLen As Long
  Dim sTemp As String * 100
    lLen = GetPrivateProfileString(Key, Section, DefVal, sTemp, 100, App.Path & "\" & INI_FILE)
    ReadINI = Mid(sTemp, 1, lLen)
End Function

Public Sub EndSync(ByVal Handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
    Mix = True
End Sub

Public Sub UnloadSync(ByVal Handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
    MixOut = True
End Sub

Public Function GetPosMin(chan As Long) As Long
Dim pos As Long
    pos = BASS_ChannelBytes2Seconds(chan, modBass.BASS_ChannelGetPosition(chan))
    GetPosMin = pos \ 60
End Function
Public Function GetPosSec(chan As Long) As Long
Dim pos As Long, Min As Long
    pos = BASS_ChannelBytes2Seconds(chan, modBass.BASS_ChannelGetPosition(chan))
    If pos > -1 Then
        Min = pos \ 60
        GetPosSec = pos - Min * 60
    Else
        GetPosSec = 0
    End If
End Function
Public Sub FadeDeckB()
    'Mp3(2).FadeComplete = True
    Master.Caption = "###########################"
End Sub

Public Sub DisableX(TheForm As Form)
    '** Description:
    '** Disable X in upper right corner of the form
    Dim lngMenu As Long
    lngMenu = GetSystemMenu(TheForm.hwnd, False)
    DeleteMenu lngMenu, 6, MF_BYPOSITION
End Sub

' RPP = Return Proper Path
Public Function RPP(ByVal fp As String) As String
    RPP = IIf(Mid(fp, Len(fp), 1) = "\", fp, fp & "\")
End Function

