VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSqlService 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SQL SERVICE MODULE"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEmpty 
      Caption         =   "Empty Current Table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2340
      TabIndex        =   10
      Top             =   6300
      Width           =   2025
   End
   Begin VB.CommandButton cmdSetRecNo 
      Caption         =   "Set RecNo"
      Height          =   315
      Left            =   1770
      TabIndex        =   9
      Top             =   6765
      Width           =   1185
   End
   Begin VB.CommandButton cmdSetSection 
      Caption         =   "Set Section"
      Height          =   315
      Left            =   135
      TabIndex        =   8
      Top             =   6750
      Width           =   1365
   End
   Begin VB.CommandButton cmddisplayList 
      Caption         =   "Display Records"
      Height          =   405
      Left            =   60
      TabIndex        =   7
      Top             =   6210
      Width           =   2205
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5910
      ItemData        =   "frmSqlService.frx":0000
      Left            =   60
      List            =   "frmSqlService.frx":0002
      TabIndex        =   6
      Top             =   270
      Width           =   2235
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9480
      TabIndex        =   4
      Top             =   6300
      Width           =   1365
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9735
      TabIndex        =   2
      Top             =   270
      Width           =   1080
   End
   Begin VB.TextBox txtSql 
      Height          =   465
      Left            =   2325
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   270
      Width           =   7365
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5490
      Left            =   2340
      TabIndex        =   0
      Top             =   780
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9684
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   " Table List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   30
      Width           =   1080
   End
   Begin VB.Label lblcurrentdatabase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "current database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2340
      TabIndex        =   3
      Top             =   30
      Width           =   1455
   End
End
Attribute VB_Name = "frmSqlService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSqlService As ADODB.Recordset
Dim tblList As ADODB.Recordset

Dim rsTRegistration As ADODB.Recordset
Dim rsTStudFinal As ADODB.Recordset
Dim rsTCourseLv As ADODB.Recordset


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDisplayList_Click()
If List1.ListIndex = -1 Then Exit Sub
On Error GoTo ErrHandlerExecute12
    If rsSqlService.state = adStateOpen Then rsSqlService.Close
    rsSqlService.CursorLocation = adUseClient
    rsSqlService.Open List1.List(List1.ListIndex), Cn, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = rsSqlService
    Exit Sub

ErrHandlerExecute12:
MsgBox "Err: " & Err.Description, vbOKOnly, "Error occured"
End Sub

Private Sub cmdEmpty_Click()
Dim tablename As String
If List1.ListIndex = -1 Then Exit Sub
On Error GoTo ErrHandlerExecute13

    If MsgBox("Empty Selected Table?", vbYesNo, "Empty Table") = vbNo Then Exit Sub
    
    tablename = List1.List(List1.ListIndex)
    Cn.Execute "DELETE FROM " & tablename & " "
    
    Set DataGrid1.DataSource = Nothing
    DataGrid1.Refresh
    Exit Sub
ErrHandlerExecute13:
MsgBox "Err: " & Err.Description, vbOKOnly, "Error occured"

End Sub

Private Sub cmdExecute_Click()
On Error GoTo ErrHandlerExecute

If rsSqlService.state = adStateOpen Then rsSqlService.Close
rsSqlService.CursorLocation = adUseClient
rsSqlService.Open txtSql.text, Cn, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rsSqlService

Exit Sub

ErrHandlerExecute:
MsgBox "Err: " & Err.Description, vbOKOnly, "Error occured"
End Sub

Private Sub cmdSetRecNo_Click()
Dim sqltxt As String
Dim SecTxt As String
Dim i As Integer

If rsTStudFinal.state = adStateOpen Then rsTStudFinal.Close
rsTStudFinal.CursorLocation = adUseClient
rsTStudFinal.Open "SELECT * FROM TStudFinal WHERE SchYr='2003-2004' AND RecNo=0", Cn, adOpenDynamic, adLockOptimistic
If rsTStudFinal.RecordCount <> 0 Then
    Do While Not rsTStudFinal.EOF
    
        cmdSetRecNo.Caption = rsTStudFinal.RecordCount - rsTStudFinal.AbsolutePosition
        
        If rsTCourseLv.state = adStateOpen Then rsTCourseLv.Close
        rsTCourseLv.CursorLocation = adUseClient
        sqltxt = "SELECT * From TCourseLv Where SCode='" & rsTStudFinal.Fields!scode & "' AND CNo=" & rsTStudFinal.Fields!cno
        rsTCourseLv.Open sqltxt, Cn, adOpenStatic, adLockOptimistic
        If rsTCourseLv.RecordCount <> 0 Then
            rsTStudFinal.Fields!RecNo = rsTCourseLv.Fields!RecNo
            rsTStudFinal.Update
        End If
        
        
    rsTStudFinal.MoveNext
    Loop
End If

End Sub

Private Sub cmdSetSection_Click()
Dim sqltxt As String
Dim SecTxt As String
Dim i As Integer

If rsTRegistration.state = adStateOpen Then rsTRegistration.Close
rsTRegistration.CursorLocation = adUseClient
rsTRegistration.Open "SELECT * FROM TRegistration WHERE SY='2003-2004'", Cn, adOpenDynamic, adLockOptimistic
If rsTRegistration.RecordCount <> 0 Then
    Do While Not rsTRegistration.EOF
        cmdSetSection.Caption = rsTRegistration.RecordCount - rsTRegistration.AbsolutePosition
        If rsTStudFinal.state = adStateOpen Then rsTStudFinal.Close
        rsTStudFinal.CursorLocation = adUseClient
        sqltxt = "SELECT TStudFinal.Sec, TStudFinal.SchYr, Count(TStudFinal.Sec) AS CountOfSec From TStudFinal Where TStudFinal.SRecNo = " & rsTRegistration.Fields!SRecNo & " GROUP BY TStudFinal.Sec, TStudFinal.SchYr HAVING TStudFinal.SchYr='2003-2004'"
        rsTStudFinal.Open sqltxt, Cn, adOpenStatic, adLockOptimistic
        If rsTStudFinal.RecordCount <> 0 Then
             i = 0
             SecTxt = rsTStudFinal.Fields!Sec
            
             'find the common subject section use by the student
             Do While Not rsTStudFinal.EOF
                If rsTStudFinal.Fields!CountOfSec > i Then
                    i = rsTStudFinal.Fields!CountOfSec
                    SecTxt = rsTStudFinal.Fields!Sec
                End If
                rsTStudFinal.MoveNext
             Loop
            rsTRegistration.Fields!Sec = SecTxt
            rsTRegistration.Update
        End If
    rsTRegistration.MoveNext
    Loop
End If

End Sub



Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    rsSqlService.Update
End Sub


Private Sub Form_Load()
    Dim i As Integer
    Dim tblcount As Integer
    
    Screen.MousePointer = vbHourglass
    
   
    
    Set rsSqlService = New ADODB.Recordset
    Set tblList = New ADODB.Recordset
    Set tblList = Cn.OpenSchema(adSchemaTables)
    
    Set rsTRegistration = New ADODB.Recordset
    Set rsTStudFinal = New ADODB.Recordset
    Set rsTCourseLv = New ADODB.Recordset
    
    'list all the table from the database
        Do While Not tblList.EOF
            If tblList.Fields(3) = "TABLE" Then
            List1.AddItem tblList.Fields(2)
            End If
            tblList.MoveNext
        Loop
        
    lblcurrentdatabase.Caption = ReadINI("Database", "Path", "cas00000.mdb")
    Screen.MousePointer = vbNormal
End Sub


Private Sub List1_DblClick()
    cmdDisplayList_Click
End Sub
