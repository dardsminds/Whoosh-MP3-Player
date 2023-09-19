Attribute VB_Name = "modPublic"
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SetWindowPos Lib "user32" (ByVal h%, ByVal hb%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
'API calls used for reading and writing of preferences
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Long) As Long

'Constants for changing the treeview
Public Const GWL_STYLE = -16&
Public Const TVM_SETBKCOLOR = 4381&
Public Const TVM_GETBKCOLOR = 4383&
Public Const TVS_HASLINES = 2&
Public Const TV_FIRST As Long = &H1100
Public Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)
Public Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Public Const TVS_CHECKBOXES = &H100
Public Const TVS_TRACKSELECT = &H200


' Win32 Declarations for DisableX
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Const INI_FILE = "Whoosh.ini"


Type BASSENGINE
    OnAirDevice As Integer
    CueDevice As Integer
    OutputQuality As Long
    MemoryBuffer As Single
    EnableOnAirDevice As Integer
    EnableCueDevice As Integer
End Type

Public AutoDJ As Boolean
Public PlSourceRow As Integer
Public IsDrag As Boolean
Public AUDIO As BASSENGINE

Public Sample1 As Long
Public Sample2 As Long
Public Sample3 As Long
Public Sample4 As Long

Public GainAdj(0 To 200) As Integer
Public CurrentCategory As String

Public IsPlaying As Boolean
Public TurnTableIsActive As Boolean

Public Player1MixNext As Boolean
Public Player2MixNext As Boolean
Public PriorityPlayer As Integer

Public GENINFO As DATABASE_INFO

Sub Main()
    
    If App.PrevInstance Then
        MsgBox "Program is already running...", vbOKOnly, "Error"
        End
    End If
    
    frmSplash.Show
    OnTop frmSplash, True
    frmSplash.SetFocus
    
    frmSplash.lblStat.Caption = "Loading database..."
    DoEvents
    
    'connect to datbase
    GENINFO.DatabaseName = ReadINI("DATABASE", "DatabasePath", "")

    OpenMDB

    If Cn.state <> adStateOpen Then
        Unload frmSplash
        MsgBox "Database not found, click ok to search for database.", vbOKOnly, "Database Error"
        FrmDatabaseDir.Show 1
        End
    End If
    
    ChDrive App.Path
    ChDir App.Path
    
    
    frmSplash.lblStat.Caption = "Reading settings..."
    DoEvents
    
    'load settings
    AUDIO.CueDevice = ReadINI("OUTPUT", "CueDevice", "1")
    AUDIO.EnableCueDevice = ReadINI("OUTPUT", "CueDeviceEnabled", "0")
    AUDIO.EnableOnAirDevice = ReadINI("OUTPUT", "DeviceEnabled", "1")
    AUDIO.MemoryBuffer = ReadINI("OUTPUT", "MemoryBuffer", "2.5")
    AUDIO.OnAirDevice = ReadINI("OUTPUT", "Device", "1")
    AUDIO.OutputQuality = ReadINI("OUTPUT", "Quality", "44100")
    
    
    frmSplash.lblStat.Caption = "Initializing BASS sound engine..."
    DoEvents
    
    'check if "BASS.DLL" is exists
    If FileExist(RPP(App.Path) & "bass.dll") = False Then
        MsgBox "Error: BASS.DLL does not exists", vbCritical, "BASS.DLL"
        End
    End If
    
    'Check that at least BASS 2.0 is loaded
    'If BASS_GetVersion < MakeLong(2, 0) Then
    '    Call MsgBox("Error: BASS version 2.0 or greater was not loaded", vbCritical, "BASS.DLL")
    '    End
    'End If
    
    'Check that "BASS_FX.dll" is exists
    If Not FileExist(RPP(App.Path) & "bass_fx.dll") Then
        Call MsgBox("BASS_FX.DLL does not exists!", vbCritical, "BASS_FX.DLL")
        End
    End If
  
    'Check that BASS_FX 2.0 was loaded
    'If BASS_FX_GetVersion <> MakeLong(2, 0) Then
     '   Call MsgBox("Error: BASS_FX version 2.0 was not loaded", vbCritical, "BASS_FX.DLL")
    '    End
    'End If
    
   
   If (BASS_Init(AUDIO.OnAirDevice, AUDIO.OutputQuality, 0, MainFrm.hwnd, 0) = 0) Then
        MsgBox "Error: Couldn't Initialize Digital Output #1", vbCritical, "Digital output"
        End
   End If
        
   If AUDIO.EnableCueDevice = 1 Then
    If (BASS_Init(AUDIO.CueDevice, AUDIO.OutputQuality, 0, MainFrm.hwnd, 0) = 0) Then
        MsgBox "Error: Couldn't Initialize Digital Output #1", vbCritical, "Digital output"
        End
    End If
   End If
        
    frmSplash.lblStat.Caption = "Loading samples..."
    DoEvents
        
   Sample1 = BASS_SampleLoad(BASSFALSE, App.Path & "\samples\sample1.mp3", 0, 0, 3, BASS_SAMPLE_OVER_POS)
   Sample2 = BASS_SampleLoad(BASSFALSE, App.Path & "\samples\sample2.wav", 0, 0, 3, BASS_SAMPLE_OVER_POS)
   Sample3 = BASS_SampleLoad(BASSFALSE, App.Path & "\samples\sample3.wav", 0, 0, 3, BASS_SAMPLE_OVER_POS)
   Sample4 = BASS_SampleLoad(BASSFALSE, App.Path & "\samples\sample4.wav", 0, 0, 3, BASS_SAMPLE_OVER_POS)
   


    'cn.

    frmSplash.lblStat.Caption = "Loading main interface..."
    DoEvents

    Load MainFrm
    
    Unload frmSplash
    
     If ReadINI("MAIN", "ShowPlaylistGenerator", vbChecked) = vbChecked Then
        PLaylistGeneratorFrm.Show 1
    End If
End Sub


Public Sub CloseProgram()
    Set Cn = Nothing
    End
End Sub


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

Public Sub PaintBackground(dfrm As Form, BgSrc As PictureBox)
Dim x As Integer
Dim y As Integer
For y = 0 To dfrm.ScaleHeight Step BgSrc.ScaleHeight
    For x = 0 To dfrm.ScaleWidth Step BgSrc.ScaleWidth
         BitBlt dfrm.hdc, x, y, BgSrc.ScaleWidth, BgSrc.ScaleHeight, BgSrc.hdc, 0, 0, vbSrcCopy
         DoEvents
    Next x
Next y
End Sub

Public Function SetIcon(cat As String) As Integer
       Select Case Trim(cat)
       Case Is = "~StationID"
           SetIcon = 5
       Case Is = "~VoiceOverBackground"
           SetIcon = 4
       Case Is = "~Jingles"
           SetIcon = 5
       Case Else
            SetIcon = 3
       End Select
End Function
