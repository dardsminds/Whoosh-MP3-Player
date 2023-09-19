VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmErrorMessage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Error Messages"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "frmErrorMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Always show this dialog box"
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   2475
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   360
      Left            =   4725
      TabIndex        =   2
      Top             =   150
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   4185
      Top             =   1185
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4605
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErrorMessage.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErrorMessage.frx":055E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtErrorMsg 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   855
      Width           =   5835
   End
   Begin VB.Label Label1 
      Caption         =   "The program generates an Error message below, please send this error to whooshsupport@mindworksoft.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   705
      TabIndex        =   1
      Top             =   135
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmErrorMessage.frx":0622
      Top             =   105
      Width           =   480
   End
End
Attribute VB_Name = "frmErrorMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tgle As Boolean

Private Sub cmdClose_Click()
    SaveINI "ErrorMessage", "Visible", Check1.Value
    
    'frmMain.StatusBar1.Panels(6).Picture = Nothing
    'frmMain.StatusBar1.Panels(6).text = ""
    Unload Me
End Sub

Private Sub Form_Load()
    Timer1.Interval = 500
    Timer1.Enabled = True
    'frmMain.StatusBar1.Panels(6).text = "System Error!"
    
    txtErrorMsg.text = "Error Number: " & ERRMSG.ErrorNumber & vbCrLf & vbCrLf & "Error Description: " & vbCrLf & ERRMSG.ErrorDescription & vbCrLf & vbCrLf & "Source: " & ERRMSG.ErrorSource
    
    Check1.Value = Val(ReadINI("ErrorMessage", "Visible", "1"))
    
    If Check1.Value = vbChecked Then
        Me.Show
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
'    tgle = Not tgle
'    If tgle = False Then
'    frmMain.StatusBar1.Panels(6).Picture = ImageList1.ListImages(1).Picture
'    Else
'    frmMain.StatusBar1.Panels(6).Picture = ImageList1.ListImages(2).Picture
'    End If
End Sub


Private Sub txtErrorMsg_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
