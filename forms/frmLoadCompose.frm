VERSION 5.00
Begin VB.Form frmLoadCompose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load playlist"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   Icon            =   "frmLoadCompose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   3705
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   345
      Left            =   1140
      TabIndex        =   1
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Playlist names"
      Height          =   4545
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3765
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "To delete selected playlist name, just press the delete key"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   150
         TabIndex        =   3
         Top             =   4050
         Width           =   3285
      End
   End
End
Attribute VB_Name = "frmLoadCompose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsplaylistname As ADODB.Recordset
Dim sqltxt As String

Private Sub cmdOK_Click()

 If List1.listcount <> 0 Then
    cmdOK.Enabled = False
    MainFrm.LoadCompose List1.text
    cmdOK.Enabled = True
 End If
 
 Unload Me
End Sub

Private Sub Form_Load()
    LoadPlaylistName
End Sub

Private Sub LoadPlaylistName()
    Set rsplaylistname = OpenRS("SELECT DISTINCT composename FROM composition ORDER BY composename")
    List1.Clear
    If rsplaylistname.RecordCount <> 0 Then
        Do While Not rsplaylistname.EOF
            List1.AddItem rsplaylistname.Fields!composename
            rsplaylistname.MoveNext
        Loop
    End If
End Sub


Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 And List1.listcount > 0 Then
        If MsgBox("Do you want to delete selected playlist name?", vbYesNo, "Delete playlist name") = vbNo Then Exit Sub
        sqltxt = "DELETE * FROM composition WHERE composename='" & List1.text & "'"
        ExecuteQuery sqltxt
        LoadPlaylistName
    
    End If
    
End Sub


