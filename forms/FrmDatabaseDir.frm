VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmDatabaseDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Directory Setup"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   ControlBox      =   0   'False
   Icon            =   "FrmDatabaseDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4125
      Top             =   -225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Database Location"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2430
      TabIndex        =   3
      Top             =   1815
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   345
      Left            =   1215
      TabIndex        =   2
      Top             =   1815
      Width           =   1185
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   345
      Left            =   3555
      TabIndex        =   1
      Top             =   1275
      Width           =   1185
   End
   Begin VB.TextBox txtPath 
      Height          =   675
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   510
      Width           =   4500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database location"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   75
      TabIndex        =   4
      Top             =   15
      Width           =   4800
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File Path:"
         Height          =   180
         Left            =   180
         TabIndex        =   5
         Top             =   255
         Width           =   795
      End
   End
End
Attribute VB_Name = "FrmDatabaseDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cntrl As Control
Dim rsdbid As ADODB.Recordset


Private Sub cmdBrowse_Click()
  ' Set CancelError is True
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  ' Set flags
  CommonDialog1.flags = cdlOFNHideReadOnly
  ' Set filters
  CommonDialog1.Filter = "Whoosh Database |data.MDB"
  ' Display the Open dialog box
  CommonDialog1.ShowOpen
  ' Display name of selected file
  txtPath.text = CommonDialog1.FileName
  If Trim(txtPath.text) <> "" Then
    cmdSave.Enabled = True
  Else
    cmdSave.Enabled = False
  End If
  
  Exit Sub
  
ErrHandler:
  'User pressed the Cancel button
  Exit Sub
End Sub

Private Sub cmdcancel_Click()
Unload Me
End
End Sub

Private Sub cmdsave_Click()
    On Error GoTo ErrHandlerSave
    GENINFO.DatabaseName = txtPath.text
    SaveINI "DATABASE", "DatabasePath", GENINFO.DatabaseName
    MsgBox "Your must restart the program in order the new settings will take effect", vbInformation, "Settings Updated!"
    End
    Exit Sub
ErrHandlerSave:
    If Err.Number = -2147217865 Then
        MsgBox "The selected database does not contain database ID information", vbOKOnly, "Database not compatible"
        End
    End If
    If Err.Number = -2147217843 Then 'invalid password
        MsgBox "Selected database has invalid or different password", vbOKOnly, "Invalid database password"
        End
    End If
    MsgBox Err.Description, vbOKOnly, Err.Number
End Sub


Private Sub Form_Load()
    Set rsdbid = New ADODB.Recordset
End Sub



