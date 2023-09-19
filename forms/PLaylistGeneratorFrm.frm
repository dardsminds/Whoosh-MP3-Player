VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form PLaylistGeneratorFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Playlist Generator"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PLaylistGeneratorFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGen 
      Caption         =   "Generate"
      Height          =   345
      Left            =   3990
      TabIndex        =   27
      Top             =   4710
      Width           =   1140
   End
   Begin VB.CheckBox chkPlayInstant 
      Caption         =   "Play immediately after generate"
      Height          =   255
      Left            =   2550
      TabIndex        =   26
      Top             =   5130
      Width           =   2655
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Show this dialog at startup"
      Height          =   255
      Left            =   90
      TabIndex        =   25
      Top             =   5130
      Width           =   2325
   End
   Begin VB.Frame Frame6 
      Caption         =   "Action"
      Height          =   645
      Left            =   3270
      TabIndex        =   21
      Top             =   2100
      Width           =   3060
      Begin VB.CheckBox chkRaplace 
         Caption         =   "Replace songs on the playlist"
         Height          =   240
         Left            =   135
         TabIndex        =   22
         Top             =   270
         Value           =   1  'Checked
         Width           =   2490
      End
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   105
      TabIndex        =   20
      Top             =   4740
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame5 
      Caption         =   "Station ID Insertions"
      ForeColor       =   &H00B87B38&
      Height          =   645
      Left            =   3270
      TabIndex        =   15
      Top             =   1380
      Width           =   3060
      Begin VB.TextBox txtInsItem 
         Height          =   315
         Left            =   2055
         TabIndex        =   17
         Text            =   "5"
         Top             =   225
         Width           =   450
      End
      Begin VB.CheckBox chkInsertID 
         Caption         =   "Insert Station ID every"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   270
         Width           =   1980
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "items."
         Height          =   240
         Left            =   2580
         TabIndex        =   18
         Top             =   270
         Width           =   510
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Playlist length"
      Enabled         =   0   'False
      ForeColor       =   &H00B87B38&
      Height          =   645
      Left            =   90
      TabIndex        =   11
      Top             =   3930
      Width           =   6270
      Begin VB.ComboBox cmbPLlength 
         Height          =   330
         ItemData        =   "PLaylistGeneratorFrm.frx":000C
         Left            =   2880
         List            =   "PLaylistGeneratorFrm.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   210
         Width           =   1425
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   2145
         TabIndex        =   13
         Text            =   "8"
         Top             =   210
         Width           =   675
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Playlist should run at least "
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   255
         Width           =   1950
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item separation"
      ForeColor       =   &H00B87B38&
      Height          =   990
      Left            =   90
      TabIndex        =   5
      Top             =   2850
      Width           =   6240
      Begin VB.CheckBox chkTimeSep 
         Caption         =   "Exclude any items played during the last"
         Height          =   240
         Left            =   150
         TabIndex        =   23
         Top             =   615
         Value           =   1  'Checked
         Width           =   3420
      End
      Begin VB.ComboBox cmbTime 
         Height          =   330
         ItemData        =   "PLaylistGeneratorFrm.frx":0033
         Left            =   4500
         List            =   "PLaylistGeneratorFrm.frx":0040
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   585
         Width           =   1185
      End
      Begin VB.TextBox txtTimeSep 
         Height          =   315
         Left            =   3735
         TabIndex        =   9
         Text            =   "60"
         Top             =   585
         Width           =   675
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2505
         TabIndex        =   7
         Text            =   "10"
         Top             =   210
         Width           =   840
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "items."
         Height          =   240
         Left            =   3480
         TabIndex        =   8
         Top             =   255
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Separate item by artist at least "
         Enabled         =   0   'False
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   255
         Width           =   2385
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rating"
      ForeColor       =   &H00B87B38&
      Height          =   1290
      Left            =   3270
      TabIndex        =   2
      Top             =   30
      Width           =   3060
      Begin VB.CheckBox chkRating 
         Caption         =   "Only items which have a rating of"
         Height          =   240
         Left            =   180
         TabIndex        =   24
         Top             =   255
         Value           =   1  'Checked
         Width           =   2820
      End
      Begin VB.ComboBox cmbRating 
         Height          =   330
         ItemData        =   "PLaylistGeneratorFrm.frx":005A
         Left            =   180
         List            =   "PLaylistGeneratorFrm.frx":007C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   690
         Width           =   1035
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " or higher (1 > 10)"
         Height          =   240
         Left            =   1380
         TabIndex        =   4
         Top             =   750
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Category"
      ForeColor       =   &H00B87B38&
      Height          =   2745
      Left            =   75
      TabIndex        =   1
      Top             =   15
      Width           =   3090
      Begin VB.ListBox lstCategory 
         BackColor       =   &H00FFFFFF&
         Height          =   2220
         ItemData        =   "PLaylistGeneratorFrm.frx":009F
         Left            =   120
         List            =   "PLaylistGeneratorFrm.frx":00A1
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   195
         Width           =   2820
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   345
      Left            =   5220
      TabIndex        =   0
      Top             =   5220
      Width           =   1065
   End
End
Attribute VB_Name = "PLaylistGeneratorFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpMUSIC() As MUSICFILE
Dim rsMusic As ADODB.Recordset
Dim rsStationID As ADODB.Recordset
Dim rsFolder As ADODB.Recordset
Dim rsMusiclist As ADODB.Recordset


Private Sub cmdClose_Click()
    Call SaveINI("MAIN", "ShowPlaylistGenerator", chkShow.value)
    Call SaveINI("MAIN", "InstantPlay", chkPlayInstant.value)
    Unload Me
End Sub

Private Sub cmdGen_Click()
    Dim i As Integer
    Dim s As String
    Dim sqltxt As String
    Dim dDiff As Long
    Dim tDiff As Long
    Dim LastPlay As String
    Dim sepMinute As Long
    Dim tmpBUFF() As Integer
    Dim MusBuff() As Long
    Dim idx As Long
        
        
    sepMinute = 5
    If cmbTime.text = "Minutes" Then sepMinute = Val(txtTimeSep.text)
    If cmbTime.text = "Hours" Then sepMinute = Val(txtTimeSep.text) * 60
    If cmbTime.text = "Days" Then sepMinute = Val(txtTimeSep.text) * 1440
    
    
    Cn.Execute "DELETE FROM tmpmusic" 'clear the temp table
    Cn.Execute "DELETE FROM tmpplaylist" 'clear the temp table

    Me.MousePointer = vbHourglass

    For i = 0 To lstCategory.listcount - 1
        If lstCategory.Selected(i) = True Then
            s = lstCategory.List(i)
            Set rsMusic = OpenRS("SELECT * FROM music WHERE category='" & s & "'")
            If rsMusic.RecordCount <> 0 Then
                Do While Not rsMusic.EOF
                    sqltxt = ""
                    LastPlay = ""
                    
                    If FileExist(rsMusic.Fields!file) = False Then GoTo JMPCONTINUE
                    
                    sqltxt = "INSERT INTO tmpmusic(`index`,`file`,`title`,`artist`,`SongStart`,`SongEnd`,`bpm`,`Len`,`Category`,`LastPlay`,`MixType`,`Rating`,`Genre`,`Album`,`SongYear`) "
                    sqltxt = sqltxt & " VALUES('" & rsMusic.Fields!Index & "','" & QuoteReplace(rsMusic.Fields!file) & " ',"
                    sqltxt = sqltxt & "'" & QuoteReplace(rsMusic.Fields!Title) & " ',"
                    sqltxt = sqltxt & "'" & QuoteReplace(rsMusic.Fields!Artist) & " ',"
                    sqltxt = sqltxt & rsMusic.Fields!SongStart & ","
                    sqltxt = sqltxt & rsMusic.Fields!SongEnd & ","
                    sqltxt = sqltxt & rsMusic.Fields!bpm & ","
                    sqltxt = sqltxt & rsMusic.Fields!Len & ","
                    sqltxt = sqltxt & "'" & rsMusic.Fields!Category & " ',"
                    sqltxt = sqltxt & "'" & rsMusic.Fields!LastPlay & " ',"
                    sqltxt = sqltxt & "'" & rsMusic.Fields!MixType & " ',"
                    sqltxt = sqltxt & rsMusic.Fields!Rating & ","
                    sqltxt = sqltxt & "'" & rsMusic.Fields!Genre & " ',"
                    sqltxt = sqltxt & "'" & QuoteReplace(rsMusic.Fields!Album) & " ',"
                    sqltxt = sqltxt & "' ')"
                    
                    LastPlay = rsMusic.Fields!LastPlay
                    
                    If chkTimeSep.value = vbChecked And Trim(LastPlay) <> "" Then

                        tDiff = DateDiff("n", TimeValue(LastPlay), time)
                        dDiff = DateDiff("d", DateValue(LastPlay), Date)
                        
                        If dDiff > 0 Then
                            tDiff = tDiff + (dDiff * 1440)
                        End If
                        
                        If tDiff > sepMinute Then
                            ExecuteQuery sqltxt
                        End If
                    Else
                        ExecuteQuery sqltxt
                    End If
                    
JMPCONTINUE:

                    rsMusic.MoveNext
                Loop
            End If
        End If
    Next i


    'randomized
    Set rsMusic = OpenRS("SELECT * FROM tmpmusic")
    If rsMusic.RecordCount <> 0 Then
        i = 0
        ReDim MusBuff(rsMusic.RecordCount) As Long
        Call RandomNumbers(tmpBUFF(), 1, rsMusic.RecordCount)
        Do While Not rsMusic.EOF
            i = i + 1
            MusBuff(i) = rsMusic.Fields!Index
            rsMusic.MoveNext
        Loop
    End If
    
    'copy the randomize music to temp playlist table
    For i = 1 To UBound(MusBuff)
           idx = MusBuff(tmpBUFF(i))
           rsMusic.MoveFirst
           rsMusic.Find "index like '" & idx & "'"
            
            sqltxt = ""
           
            sqltxt = "INSERT INTO tmpplaylist(`index`,`file`,`title`,`artist`,`SongStart`,`SongEnd`,`bpm`,`Len`,`Category`,`LastPlay`,`MixType`,`Rating`,`Genre`,`Album`,`SongYear`) "
            sqltxt = sqltxt & " VALUES('" & rsMusic.Fields!Index & "','" & QuoteReplace(rsMusic.Fields!file) & " ',"
            sqltxt = sqltxt & "'" & QuoteReplace(rsMusic.Fields!Title) & " ',"
            sqltxt = sqltxt & "'" & QuoteReplace(rsMusic.Fields!Artist) & " ',"
            sqltxt = sqltxt & rsMusic.Fields!SongStart & ","
            sqltxt = sqltxt & rsMusic.Fields!SongEnd & ","
            sqltxt = sqltxt & rsMusic.Fields!bpm & ","
            sqltxt = sqltxt & rsMusic.Fields!Len & ","
            sqltxt = sqltxt & "'" & rsMusic.Fields!Category & " ',"
            sqltxt = sqltxt & "'" & rsMusic.Fields!LastPlay & " ',"
            sqltxt = sqltxt & "'" & rsMusic.Fields!MixType & " ',"
            sqltxt = sqltxt & rsMusic.Fields!Rating & ","
            sqltxt = sqltxt & "'" & rsMusic.Fields!Genre & " ',"
            sqltxt = sqltxt & "'" & QuoteReplace(rsMusic.Fields!Album) & " ',"
            sqltxt = sqltxt & "' ')"
            ExecuteQuery sqltxt
           
    Next i
    
    'insert Station ID
'    If chkInsertID.value = vbChecked Then
'    Set rsStationID = OpenRS("SELECT * FROM music WHERE Category='~StationID'")
'    'rsStationID
'
'
'
'    End If
    
    'list the song on the playlist
    MainFrm.ListView1.ListItems.Clear
    Set rsMusic = OpenRS("SELECT * FROM tmpplaylist")
    Do While Not rsMusic.EOF
           Set lst = MainFrm.ListView1.ListItems.Add()
           lst.text = "-"
           lst.SubItems(1) = rsMusic.Fields!Title
           lst.SubItems(2) = rsMusic.Fields!Artist
           lst.SubItems(3) = Format(rsMusic.Fields!bpm, "000")
           lst.SubItems(4) = rsMusic.Fields!Index
           lst.SubItems(5) = rsMusic.Fields!Genre
           lst.SubItems(7) = rsMusic.Fields!Rating
           lst.SubItems(8) = rsMusic.Fields!LastPlay
           lst.SubItems(9) = rsMusic.Fields!Len
           lst.SmallIcon = SetIcon(rsMusic.Fields!Category)
           rsMusic.MoveNext
    Loop
    
    MainFrm.UpdateTimePlay
    
    Me.MousePointer = vbNormal
    
End Sub

Private Sub RandomNumbers(RandomBuff() As Integer, lowerNum As Integer, UpperNum As Integer)
  Dim k, Range As Integer
  Dim i, nRandom As Integer
  Range = UpperNum - lowerNum + 1
  ReDim Preserve RandomBuff(1 To Range) As Integer
  For k = 1 To Range
    Randomize
    nRandom = GetRandomNumber(lowerNum, UpperNum)
      For i = 1 To k - 1
        Do Until RandomBuff(i) <> nRandom
          If RandomBuff(i) = nRandom Then
            nRandom = GetRandomNumber(lowerNum, UpperNum)
            i = 1
          End If
        Loop
      Next i
    RandomBuff(k) = nRandom
  Next k
End Sub

Private Function GetRandomNumber(lowerrange As Integer, upperrange As Integer)
    GetRandomNumber = Int((upperrange - lowerrange + 1) * Rnd + lowerrange)
End Function

Private Sub Form_Load()
    Dim i As Integer
    Call LoadGenre
    
    cmbTime.ListIndex = 0
    cmbPLlength.ListIndex = 0
    cmbRating.ListIndex = 9
    
    'cmbGenre.ListIndex = 1
    chkShow.value = ReadINI("MAIN", "ShowPlaylistGenerator", vbChecked)
    chkPlayInstant.value = ReadINI("MAIN", "InstantPlay", vbChecked)

    Set rsFolder = OpenRS("SELECT * FROM folder ORDER BY folder")
    If rsFolder.RecordCount <> 0 Then
        Do While Not rsFolder.EOF
        lstCategory.AddItem rsFolder.Fields!folder
        rsFolder.MoveNext
        Loop
        lstCategory.ListIndex = 1
    End If
End Sub

Private Function IsPlayable(Index As Integer, timeSeparation As Long) As Boolean
    Dim LastPlay As String
    Dim dDiff As Long
    Dim tDiff As Long
    
    LastPlay = MUSIC(Index).LastPlay
    If Trim(LastPlay) = "" Then
        IsPlayable = True
        Exit Function
    End If
    
    tDiff = DateDiff("n", TimeValue(LastPlay), time)
    dDiff = DateDiff("d", DateValue(LastPlay), Date)
       
    If dDiff > 0 Then
        tDiff = tDiff + (dDiff * 1440)
    End If
        
        
    If tDiff < timeSeparation Then
        IsPlayable = False
    Else
        IsPlayable = True
    End If

    'MsgBox DateValue(LastPlay) & "-" & tDiff & "-" & timeSeparation & "-" & Date & "-" & IsPlayable
    
    Exit Function

End Function

