VERSION 5.00
Begin VB.Form frmSaveCompose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save current playlist"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmSaveCompose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Height          =   345
      Left            =   1560
      TabIndex        =   1
      Top             =   1290
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Playlist information"
      Height          =   1155
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   4425
      Begin VB.TextBox txtDjname 
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   660
         Width           =   3210
      End
      Begin VB.TextBox txtComposename 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   270
         Width           =   3210
      End
      Begin VB.Label Label2 
         Caption         =   "DJ name"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   675
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Playlist name"
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   315
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmSaveCompose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCompose As ADODB.Recordset

Private Sub cmdOK_Click()
 Dim i As Integer
 Dim lidx As Long
  
  
 If Trim(txtComposename.text) = "" Then
    Me.Hide
    MsgBox "Please enter playlist name", vbOKOnly, "Save playlist"
    Exit Sub
    Me.Show
    Exit Sub
 End If
If Trim(txtDjname.text) = "" Then
    Me.Hide
    MsgBox "Please enter DJ name", vbOKOnly, "Save playlist"
    Exit Sub
    Me.Show
    Exit Sub
 End If
  
  
  
 txtComposename.Enabled = False
 txtDjname.Enabled = False
 cmdOK.Enabled = False
 
 With MainFrm
    For i = 1 To .ListView1.ListItems.count
        lidx = .ListView1.ListItems.Item(i).SubItems(4)
        SaveCompose lidx, txtComposename.text, txtDjname.text
    Next i
    
 End With

 txtComposename.Enabled = True
 txtDjname.Enabled = True
 cmdOK.Enabled = True
 
 Unload Me
End Sub


Private Sub SaveCompose(idx As Long, composename As String, djname As String)
    Dim sqltxt As String
    Set rsCompose = OpenRS("SELECT * FROM music WHERE index=" & idx)
    If rsCompose.RecordCount <> 0 Then
        sqltxt = "INSERT INTO composition(musicindex,len,SongStart,SongEnd,MixType,composename,dj) "
        sqltxt = sqltxt & " VALUES(" & idx & "," & rsCompose.Fields!Len & "," & rsCompose.Fields!SongStart & "," & rsCompose.Fields!SongEnd & ",'" & rsCompose.Fields!MixType & "','" & composename & "','" & djname & "')"
        ExecuteQuery sqltxt
    End If
End Sub
