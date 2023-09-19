VERSION 5.00
Begin VB.Form frmAddStation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Internet Radio Station"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmAddStation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "Done"
      Height          =   345
      Left            =   1410
      TabIndex        =   1
      Top             =   1545
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Station Information"
      Height          =   1485
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   4425
      Begin VB.TextBox txtGenre 
         Height          =   315
         Left            =   1140
         TabIndex        =   7
         Top             =   1035
         Width           =   2490
      End
      Begin VB.TextBox txtHostAddress 
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   660
         Width           =   3210
      End
      Begin VB.TextBox txtStationName 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   270
         Width           =   3210
      End
      Begin VB.Label Label3 
         Caption         =   "Genre:"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label Label2 
         Caption         =   "Host:"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   675
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Station name:"
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   315
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmAddStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
 Dim stanum As Integer
  
 If Trim(txtStationName.text) = "" Then
    GoTo DONEADDING
 End If
 
 stanum = UBound(STATION)
 ReDim Preserve STATION(stanum + 1) As IRADIO
 STATION(stanum + 1).StationName = txtStationName.text
 STATION(stanum + 1).Address = txtHostAddress.text
 STATION(stanum + 1).Category = txtGenre.text
 SaveStation
 Master.RefreshStationList
 
DONEADDING:
 
 Unload Me
End Sub
