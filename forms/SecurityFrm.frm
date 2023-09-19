VERSION 5.00
Begin VB.Form SecurityFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serial Number"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   Icon            =   "SecurityFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2925
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   3555
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   750
         Picture         =   "SecurityFrm.frx":0442
         ScaleHeight     =   375
         ScaleWidth      =   1815
         TabIndex        =   4
         Top             =   210
         Width           =   1845
      End
      Begin VB.TextBox txtSerial 
         Height          =   345
         Left            =   150
         TabIndex        =   2
         Top             =   1530
         Width           =   3255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "or contact: DARIO L. MINDORO at +639207874658 , email: dards@mindworksoft.com"
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   2250
         Width           =   3285
      End
      Begin VB.Label Label3 
         Caption         =   "http://www.Mindworksoft.com/whoosh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   360
         TabIndex        =   6
         Top             =   1230
         Width           =   2835
      End
      Begin VB.Label Label2 
         Caption         =   "for more info and program support just visit the website mentioned above."
         Height          =   435
         Left            =   150
         TabIndex        =   5
         Top             =   1890
         Width           =   3285
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Please enter the correct serial number, this was sent to your email during  downloading of this program from "
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   3195
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   1110
      TabIndex        =   0
      Top             =   3000
      Width           =   1335
   End
End
Attribute VB_Name = "SecurityFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sn As String

Private Sub cmdOk_Click()
sn = GetSetting("explore", "settings", "code", "04011972")
If txtSerial.text <> sn Then
    MsgBox "Invalid serial code!!!", vbCritical And vbOKOnly, "Error!!!"
    End
Else
    Call SaveSetting("explore", "settings", "code", "04011972-111")
    Unload Me
    Load Master
End If
End Sub

Private Sub Form_Load()
sn = GetSetting("explore", "settings", "code", "")
If sn = "04011972-111" Then
    Unload Me
    Load Master
End If
End Sub
