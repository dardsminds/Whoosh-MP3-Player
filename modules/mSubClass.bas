Attribute VB_Name = "mSubClass"
Public oldProc As Long
Private Declare Function CallWindowProcA Lib "user32" ( _
    ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    'wParam - the number of the hotkey, its identification.
    'lParam - HiWord is the Modifiere e.g. Shift, Ctrl, Alt
    'lParam . LoWord is the KeyCode, it is the same Code found in the Objectbrowser (F2)
    'under KeyCode
    'but I think you need only the number (identifier) of the hotkey, given in wParam.
'    Debug.Print wParam, lParam
    WndProc = 0
    If uMsg = WM_HOTKEY Then
        'The Hotkey message
        If wParam = 1 Then
            MixerFrm.MixNow
        End If
    Else
        'All other messages to the old Windowprocedure
        WndProc = CallWindowProcA(oldProc, hwnd, uMsg, wParam, lParam)
    End If
End Function

