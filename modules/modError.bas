Attribute VB_Name = "modError"
Type ERRORINFO
    ErrorDescription As String
    ErrorNumber As String
    ErrorSource As String
End Type

Public ERRMSG As ERRORINFO


Public Sub ErrorMsg(ErrorNum As String, ErrorDesc As String, ErrorSource As String, visible As Boolean)
    ERRMSG.ErrorDescription = ErrorDesc
    ERRMSG.ErrorNumber = ErrorNum
    ERRMSG.ErrorSource = ErrorSource
    If visible = True Then
        frmErrorMessage.Show
    Else
        Load frmErrorMessage
    End If
End Sub
