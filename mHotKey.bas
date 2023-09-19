Attribute VB_Name = "mHotKey"
Option Explicit

Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long

Public Enum ModConst
    MOD_ALT = &H1
    MOD_CONTROL = &H2
    MOD_SHIFT = &H4
End Enum
    
Public Const WM_HOTKEY = &H312
Private m_hkCount As Long

Function HotKeyActivate(ByVal hwnd As Long, Modifier As ModConst, Optional KeyCode As Integer) As Long
    m_hkCount = m_hkCount + 1
    ' 0 for no success, otherwise success
    HotKeyActivate = RegisterHotKey(hwnd, m_hkCount, Modifier, KeyCode)
End Function

Function HotKeyDeactivate(ByVal hwnd As Long)
    Dim i As Integer
    For i = 1 To m_hkCount
        UnregisterHotKey hwnd, i
    Next i
    m_hkCount = 0
End Function
