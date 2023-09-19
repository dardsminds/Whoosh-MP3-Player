Attribute VB_Name = "ADO_Module"
'Required add reference to Microsoft ActiveX Data Objects library 2.5
Option Explicit

Type DATABASE_INFO
    DatabaseName As String
    ServerIP As String
    ServerUID As String
    ServerPWD As String
    driver As String
    ServerPort As String
    connected As Boolean
End Type



Public Cn As ADODB.Connection
Public rst As ADODB.Recordset

Public Function OpenMDB()
    Dim strCon As String
    Dim strBuffer As String
    On Error GoTo ConnectError
    strBuffer = GENINFO.DatabaseName
    strCon = strBuffer & ";Jet OLEDB:Database Password=ariand"
    Set Cn = New ADODB.Connection
    Cn.ConnectionString = "Provider=MSDAtaShape;data provider=Microsoft.jet.oledb.4.0;Data Source=" & strCon
    Cn.Open
    GENINFO.connected = True
    Exit Function
ConnectError:
    GENINFO.connected = False
End Function


Public Sub ExecuteQuery(qstring As String)
    Cn.Execute qstring
End Sub

Public Function OpenRS(strSql As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenDynamic
        If IsNull(Cn) = False Then
            rs.Open strSql, Cn, adOpenKeyset, adLockOptimistic
        End If
        Set OpenRS = rs
End Function

Public Function QuoteReplace(s As String) As String
Dim tmpstr As String
    'find if the string contains qoutes
    If InStr(s, "'") Then
        tmpstr = Replace(s, "'", "''")
        
        
        
        QuoteReplace = tmpstr
'    ElseIf InStr(s, "\") Then
'        tmpstr = Replace(s, "\", "\\")
'        QuoteReplace = tmpstr
    Else
        QuoteReplace = s
    End If
End Function
Public Function ClsSql(str As String) As String
    Dim tmpstr As String
    Dim s As String
    s = str
    tmpstr = Replace(s, "'", "''")
    'tmpstr = Replace(s, "\", "\\")
    ClsSql = tmpstr
End Function

