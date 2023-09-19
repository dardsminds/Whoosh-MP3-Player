Attribute VB_Name = "modNetRadio"
'///////////////////////////////////////////////////////////////
' modNetRadio.bas - Copyright (c) 2002
'                              JOBnik! [Arthur Aminov, ISRAEL]
'                              e-mail: jobnik2k@hotmail.com
'
' BASS Internet radio example
'
' Originally Translated from: - netradio.c - Example of Ian Luck
'///////////////////////////////////////////////////////////////

Public WriteFile As clsFileIo
Dim FileIsOpen As Boolean
Dim DownloadStarted As Boolean

Public DlOutput As String
Public DoDownload As Boolean
Dim SongNameUpdate As Boolean

Public GotHeader As Boolean

' update stream title from metadata
Sub DoMeta(ByVal meta As Long)
    Dim P As String
    If meta = 0 Then Exit Sub
    If ((Mid(VBStrFromAnsiPtr(meta), 1, 13) = "StreamTitle='")) Then
        GotHeader = False
        DownloadStarted = False
        P = Mid(VBStrFromAnsiPtr(meta), 14)
        Master.lblRadioStatus.Caption = Mid(P, 1, InStr(P, ";") - 2)
        DlOutput = App.Path & "\" & RemoveSpecialChar(Mid(P, 1, InStr(P, ";") - 2)) & ".mp3"
    End If
End Sub

Sub MetaSync(ByVal Handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
    Call DoMeta(data)
End Sub

' The following functions where added by Peter Hebels
Public Sub UpdateFileIo()
    Set WriteFile = New clsFileIo
End Sub

Public Sub SUBDOWNLOADPROC(ByVal buffer As Long, ByVal length As Long, ByVal user As Long)

Exit Sub
If DoDownload = False Then
    DownloadStarted = False
    'WriteFile.CloseFile
    Exit Sub
End If

If DlOutput = "" Then Exit Sub

If DownloadStarted = False Then
    DownloadStarted = True
    WriteFile.CloseFile
    If WriteFile.OpenFile(DlOutput) = True Then
        SongNameUpdate = False
    Else
        SongNameUpdate = True
        GotHeader = False
    End If
End If

If SongNameUpdate = False Then
    If length <> 0 Then
        WriteFile.WriteBytes buffer, length
    Else
        WriteFile.CloseFile
        GotHeader = False
    End If
Else
    DownloadStarted = False
    WriteFile.CloseFile
    GotHeader = False
End If

End Sub

Public Function RemoveSpecialChar(StrFileName As String)
On Error Resume Next
Dim SpecialChar As Boolean
Dim SelChar As String

For i = 1 To Len(StrFileName)
    SelChar = Mid(StrFileName, i, 1)
    SpecialChar = InStr(":/\?*|<>" & Chr$(34), SelChar) > 0
    
    If SpecialChar = False Then
        OutFileName = OutFileName & SelChar
        SpecialChar = False
    Else
        OutFileName = OutFileName
        SpecialChar = False
    End If
Next i
End Function

Public Sub ReadStation()
    Dim bSize As Long
    Open App.Path & "\radio.sta" For Binary As #2
    Get #2, 1, bSize          ' Read the number of items.
    ReDim STATION(0 To bSize) As IRADIO
    Get #2, , STATION()
    Close #2
End Sub

Public Sub SaveStation()
    Open App.Path & "\radio.sta" For Binary As #1
    Put #1, 1, CLng(UBound(STATION))
    Put #1, , STATION()
    Close #1
End Sub

