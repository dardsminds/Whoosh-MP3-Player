VERSION 5.00
Begin VB.UserControl GurhanButton 
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2145
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   79
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   143
   Tag             =   "210601"
   Begin VB.PictureBox PICT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1200
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   705
      Top             =   0
   End
End
Attribute VB_Name = "GurhanButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'CREDITS:
'I benefited from a lot of people but I can not even remember their names.
'However, I remember Mr.Klaus H. Probst regarding the DrawEdge API, and
'Carles P.V. regarding ShowBorderOnFocus.
'13 July 2001
'Gurhan KARTAL
'http://gurhan.kartal.org (nothing much there :)
'gurhan@kartal.org

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As textparametreleri) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function ShellExecute _
   Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type textparametreleri
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Public Enum GB_PIC_POSITION
    gbTOP = 0
    gbLEFT = 1
    gbRIGHT = 2
    gbBOTTOM = 3
End Enum
Public Enum GB_PIC_SIZE
    size16x16 = 0
    size32x32 = 1
    sizeDefault = 2
    sizeCustom = 3
End Enum

Private mvarClientRect As RECT
Private mvarPictureRect As RECT
Private mvarCaptionRect As RECT
Dim mvarTempRect As RECT
Dim g_FocusRect As RECT
Dim alan As RECT
Dim g_TextRectUp As RECT, g_TextRectDown As RECT

Dim m_PictureOriginal As Picture
Dim m_PictureHover As Picture
Dim m_Caption As String
Dim m_PicturePosition As GB_PIC_POSITION
Dim m_Picture As Picture
Dim m_PictureWidth As Long
Dim m_PictureHeight As Long
Dim m_PictureSize As GB_PIC_SIZE
Dim mvarDrawTextParams As textparametreleri
Dim g_HasFocus As Boolean
Dim gb_MOUSE_IS_DOWN As Boolean, gb_MOUSE_IS_INSIDE As Boolean
Dim gbbBUTTON As Integer, g_Shift As Integer, g_X As Single, g_Y As Single
Dim gbKEY_PRESSED As Boolean
Dim m_URL As String
Dim m_BorderEdged As Boolean
Dim m_Raised As Boolean
Dim m_ShowBorderOnFocus As Boolean
Dim m_ShowFocusRect As Boolean

Dim WithEvents g_Font As StdFont
Attribute g_Font.VB_VarHelpID = -1

Const m_def_URL = ""
Const m_def_BorderEdged = 0
Const m_def_Raised = 0
Const m_def_ShowBorderOnFocus = True
Const m_def_ShowFocusRect = True
Const SW_SHOW = 1
Const mvarPadding As Long = 4
Const g_Light = &H80000016
Const g_Shadow = &H80000010
Const g_HighLight = &H80000014
Const g_DarkShadow = &H80000015

Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseIn(Shift As Integer)
Event MouseOut(Shift As Integer)
'**********************************************************************************

Private Sub UserControl_InitProperties()
    PICT.BackColor = Ambient.BackColor
    m_ShowBorderOnFocus = m_def_ShowBorderOnFocus
    m_ShowFocusRect = m_def_ShowFocusRect
    Set UserControl.Font = Ambient.Font
    Set g_Font = Ambient.Font
    m_Caption = Ambient.DisplayName
    m_PicturePosition = 1
    m_PictureWidth = 32
    m_PictureHeight = 32
    m_PictureSize = 1
    Set m_PictureHover = LoadPicture("")
    Set m_PictureOriginal = LoadPicture("")
    m_Raised = m_def_Raised
    m_BorderEdged = m_def_BorderEdged
    m_URL = m_def_URL
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    PICT.BackColor = UserControl.BackColor
    m_ShowFocusRect = PropBag.ReadProperty("ShowFocusRect", m_def_ShowFocusRect)
    m_ShowBorderOnFocus = PropBag.ReadProperty("ShowBorderOnFocus", m_def_ShowBorderOnFocus)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_PicturePosition = PropBag.ReadProperty("PicturePosition", 1)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_PictureWidth = PropBag.ReadProperty("PictureWidth", 32)
    m_PictureHeight = PropBag.ReadProperty("PictureHeight", 32)
    m_PictureSize = PropBag.ReadProperty("PictureSize", 1)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set m_PictureHover = PropBag.ReadProperty("PictureHover", Nothing)
    Set m_PictureOriginal = PropBag.ReadProperty("Picture", Nothing)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", Verdadero)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_Raised = PropBag.ReadProperty("Raised", m_def_Raised)
    m_BorderEdged = PropBag.ReadProperty("BorderEdged", m_def_BorderEdged)
    m_URL = PropBag.ReadProperty("URL", m_def_URL)
Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("PicturePosition", m_PicturePosition, 1)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("PictureWidth", m_PictureWidth, 32)
    Call PropBag.WriteProperty("PictureHeight", m_PictureHeight, 32)
    Call PropBag.WriteProperty("PictureSize", m_PictureSize, 1)
    Call PropBag.WriteProperty("PictureHover", m_PictureHover, Nothing)
    
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, Verdadero)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("ShowBorderOnFocus", m_ShowBorderOnFocus, m_def_ShowBorderOnFocus)
    Call PropBag.WriteProperty("ShowFocusRect", m_ShowFocusRect, m_def_ShowFocusRect)
 
    Call PropBag.WriteProperty("Raised", m_Raised, m_def_Raised)
    Call PropBag.WriteProperty("BorderEdged", m_BorderEdged, m_def_BorderEdged)
    Call PropBag.WriteProperty("URL", m_URL, m_def_URL)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
 End Sub
Private Sub CalcRECTs()
    Dim picWidth, picHeight, capWidth, capHeight As Long
    alan.Left = 0
    alan.Top = 0
    alan.Right = UserControl.ScaleWidth - 1
    alan.Bottom = UserControl.ScaleHeight - 1
    
    With mvarClientRect
     .Left = alan.Left + mvarPadding
     .Top = alan.Top + mvarPadding
     .Right = alan.Right - mvarPadding + 1
     .Bottom = alan.Bottom - mvarPadding + 1
    End With
    
    If m_Picture Is Nothing Then
        With mvarCaptionRect
           .Left = mvarClientRect.Left
           .Top = mvarClientRect.Top
           .Right = mvarClientRect.Right
           .Bottom = mvarClientRect.Bottom
        End With
        CalculateCaptionRect
    Else
        If m_Caption = "" Then
         With mvarPictureRect
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - m_PictureWidth) \ 2) + mvarClientRect.Left
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - m_PictureHeight) \ 2) + mvarClientRect.Top
            .Right = mvarPictureRect.Left + m_PictureWidth
            .Bottom = mvarPictureRect.Top + m_PictureHeight
         End With
            Exit Sub
        End If
        With mvarCaptionRect
        .Left = mvarClientRect.Left
        .Top = mvarClientRect.Top
        .Right = mvarClientRect.Right
        .Bottom = mvarClientRect.Bottom
        End With
        CalculateCaptionRect
        picWidth = m_PictureWidth
        picHeight = m_PictureHeight
        capWidth = mvarCaptionRect.Right - mvarCaptionRect.Left
        capHeight = mvarCaptionRect.Bottom - mvarCaptionRect.Top
        Select Case m_PicturePosition
        Case gbLEFT
        With mvarPictureRect
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - picHeight) \ 2) + mvarClientRect.Top
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - (picWidth + mvarPadding + capWidth)) \ 2) + mvarClientRect.Left
            .Bottom = mvarPictureRect.Top + picHeight
            .Right = mvarPictureRect.Left + picWidth
        End With
        With mvarCaptionRect
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - capHeight) \ 2) + mvarClientRect.Top
            .Left = mvarPictureRect.Right + mvarPadding
            .Bottom = mvarCaptionRect.Top + capHeight
            .Right = mvarCaptionRect.Left + capWidth
        End With
        
        Case gbRIGHT
        With mvarCaptionRect
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - capHeight) \ 2) + mvarClientRect.Top
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - (picWidth + mvarPadding + capWidth)) \ 2) + mvarClientRect.Left
            .Bottom = mvarCaptionRect.Top + capHeight
            .Right = mvarCaptionRect.Left + capWidth
        End With
        With mvarPictureRect
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - picHeight) \ 2) + mvarClientRect.Top
            .Left = mvarCaptionRect.Right + mvarPadding
            .Bottom = mvarPictureRect.Top + picHeight
            .Right = mvarPictureRect.Left + picWidth
        End With
        Case gbTOP
        With mvarPictureRect
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - (picHeight + mvarPadding + capHeight)) \ 2) + mvarClientRect.Top
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - picWidth) \ 2) + mvarClientRect.Left
            .Bottom = mvarPictureRect.Top + picHeight
            .Right = mvarPictureRect.Left + picWidth
        End With
        With mvarCaptionRect
            .Top = mvarPictureRect.Bottom + mvarPadding
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - capWidth) \ 2) + mvarClientRect.Left
            .Bottom = mvarCaptionRect.Top + capHeight
            .Right = mvarCaptionRect.Left + capWidth
        End With
        Case gbBOTTOM
        With mvarCaptionRect
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - (picHeight + mvarPadding + capHeight)) \ 2) + mvarClientRect.Top
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - capWidth) \ 2) + mvarClientRect.Left
            .Bottom = mvarCaptionRect.Top + capHeight
            .Right = mvarCaptionRect.Left + capWidth
        End With
        With mvarPictureRect
            .Top = mvarCaptionRect.Bottom + mvarPadding
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - picWidth) \ 2) + mvarClientRect.Left
            .Bottom = mvarPictureRect.Top + picHeight
            .Right = mvarPictureRect.Left + picWidth
        End With
        End Select
    End If
End Sub

Private Sub UserControl_Initialize()
    Set g_Font = New StdFont
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If Not Me.Enabled Then Exit Sub
    If KeyAscii = 13 Or KeyAscii = 27 Then
        RaiseEvent Click
        GoToURL
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    Refresh
End Sub

Private Sub UserControl_EnterFocus()
    g_HasFocus = True
    Refresh
End Sub

Private Sub UserControl_ExitFocus()
    g_HasFocus = False
    gb_MOUSE_IS_DOWN = False
    Refresh
End Sub

Private Sub UserControl_Resize()
    If ScaleWidth < 10 Then UserControl.Width = 150
    If ScaleHeight < 10 Then UserControl.Height = 150
    g_FocusRect.Left = 3
    g_FocusRect.Right = ScaleWidth - 3
    g_FocusRect.Top = 3
    g_FocusRect.Bottom = ScaleHeight - 3
    Refresh
End Sub

Public Sub Refresh()
    AutoRedraw = True
    UserControl.Cls
    CalcRECTs
    DrawPicture
    If g_HasFocus And m_ShowFocusRect Then DrawFocusRect hdc, g_FocusRect
    DrawCaption
    Draw3DEffect
    AutoRedraw = False
End Sub

Private Sub UserControl_DblClick()
    SetCapture hwnd
    UserControl_MouseDown gbbBUTTON, g_Shift, g_X, g_Y
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not gbKEY_PRESSED Then
        Select Case KeyCode
            Case vbKeyReturn
                RaiseEvent Click
                GoToURL
            Case vbKeySpace
                gb_MOUSE_IS_DOWN = True
                Refresh
                RaiseEvent Click
                GoToURL
        End Select
        gbKEY_PRESSED = True
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        gb_MOUSE_IS_DOWN = False
        Refresh
    End If
    gbKEY_PRESSED = False
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    gbbBUTTON = Button
    g_Shift = Shift
    g_X = x
    g_Y = y
    If Button <> vbRightButton Then
        gb_MOUSE_IS_DOWN = True
        Refresh
    End If
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (x >= 0 And y >= 0) And (x < ScaleWidth And y < ScaleHeight) Then
        If gb_MOUSE_IS_INSIDE = False Then
            OverTimer.Enabled = True
            gb_MOUSE_IS_INSIDE = True
            If Not m_PictureHover Is Nothing Then
                Set m_Picture = m_PictureHover
            End If
            RaiseEvent MouseIn(Shift)
            Refresh
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    gb_MOUSE_IS_DOWN = False
    If Button <> vbRightButton Then
        Refresh
        If (x >= 0 And y >= 0) And (x < ScaleWidth And y < ScaleHeight) Then
            RaiseEvent Click
            GoToURL
        End If
    End If
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    Refresh
End Property
Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = g_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    With g_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
    End With
    PropertyChanged "Font"
End Property

Private Sub g_Font_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = g_Font
    Refresh
End Sub

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
Public Property Get ShowBorderOnFocus() As Boolean
    ShowBorderOnFocus = m_ShowBorderOnFocus
End Property

Public Property Let ShowBorderOnFocus(ByVal New_ShowBorderOnFocus As Boolean)
    m_ShowBorderOnFocus = New_ShowBorderOnFocus
    PropertyChanged "ShowBorderOnFocus"
    Refresh
End Property

Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_ShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)
    m_ShowFocusRect = New_ShowFocusRect
    PropertyChanged "ShowFocusRect"
    Refresh
End Property
             
Private Sub Draw3DEffect()
    If Not Ambient.UserMode Then
         Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), g_Shadow, B
         Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), g_HighLight, B
    End If
    Select Case BorderEdged
    Case Is = False
        If gb_MOUSE_IS_DOWN Then
            Line (1, 1)-(ScaleWidth - 1, ScaleHeight - 1), g_Shadow, B
            Line (0, 0)-(ScaleWidth - 2, ScaleHeight - 2), g_Light, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), g_DarkShadow, B
            Line (-1, -1)-(ScaleWidth - 1, ScaleHeight - 1), g_HighLight, B
        End If
        If Not gb_MOUSE_IS_DOWN And gb_MOUSE_IS_INSIDE Then
            Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), g_Shadow, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), g_HighLight, B
        End If
        
        If Not gb_MOUSE_IS_DOWN And Not gb_MOUSE_IS_INSIDE And Raised Then
            Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), g_Shadow, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), g_HighLight, B
        End If
         'BORDER
          If (g_HasFocus And m_ShowBorderOnFocus And Raised And Not gb_MOUSE_IS_DOWN) Or Extender.Default Then
            Line (0, 0)-(ScaleWidth - 2, ScaleHeight - 2), g_Shadow, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - g_3DInc - 0), g_HighLight, B
            Line (-1, -1)-(ScaleWidth - 1, ScaleHeight - 1), g_DarkShadow, B
         End If
         
    Case Is = True
            Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), g_Shadow, B
            Line (1, 1)-(ScaleWidth - 2, ScaleHeight - 2), g_HighLight, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), g_DarkShadow, B
            Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), g_HighLight, B
            Line (0, 0)-(ScaleWidth - 2, ScaleHeight - 2), g_Shadow, B
    
        If gb_MOUSE_IS_DOWN Then
            Line (1, 1)-(ScaleWidth - 1, ScaleHeight - 1), g_Shadow, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), g_DarkShadow, B
            Line (-1, -1)-(ScaleWidth - 1, ScaleHeight - 1), g_HighLight, B
            Line (1.5, 1.5)-(ScaleWidth - 2, ScaleHeight - 2), g_DarkShadow, B '
        End If
        
        If Not gb_MOUSE_IS_DOWN And (gb_MOUSE_IS_INSIDE Or g_HasFocus) Then
            Line (2, 2)-(ScaleWidth - 4, 2), g_HighLight
            Line (2, 2)-(2, ScaleHeight - 3), g_HighLight
            Line (0, 0)-(ScaleWidth - 3, ScaleHeight - 3), g_DarkShadow, B
        End If
    End Select
End Sub

Private Sub OverTimer_Timer()
    Dim P As POINTAPI
    GetCursorPos P
    If hwnd <> WindowFromPoint(P.x, P.y) Then
        OverTimer.Enabled = False
        gb_MOUSE_IS_INSIDE = False
        Set m_Picture = m_PictureOriginal
        RaiseEvent MouseOut(g_Shift)
        Refresh
        If gb_MOUSE_IS_DOWN = True Then
            gb_MOUSE_IS_DOWN = False
            Refresh
            gb_MOUSE_IS_DOWN = True
        End If
    End If
End Sub

Public Property Get Raised() As Boolean
    Raised = m_Raised
End Property

Public Property Let Raised(ByVal New_Raised As Boolean)
    m_Raised = New_Raised
    PropertyChanged "Raised"
End Property

Public Property Get BorderEdged() As Boolean
    BorderEdged = m_BorderEdged
End Property

Public Property Let BorderEdged(ByVal New_BorderEdged As Boolean)
    m_BorderEdged = New_BorderEdged
    PropertyChanged "BorderEdged"
    Refresh
End Property

Public Sub GoToURL()

    If Left(m_URL, 7) = "mailto:" Then
        Navigate UserControl.Parent, m_URL
        Exit Sub
    End If
        If Not m_URL = "" Then UserControl.Hyperlink.NavigateTo m_URL
End Sub
Private Sub Navigate(frm As Form, ByVal WebPageURL As String)
Dim hBrowse As Long
hBrowse = ShellExecute(frm.hwnd, "open", WebPageURL, "", "", 1)
End Sub
Public Property Get URL() As String
    URL = m_URL
End Property

Public Property Let URL(ByVal New_URL As String)
    m_URL = New_URL
    PropertyChanged "URL"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Refresh
End Property
Public Property Get PicturePosition() As GB_PIC_POSITION
    PicturePosition = m_PicturePosition
End Property
Public Property Let PicturePosition(ByVal New_PicturePosition As GB_PIC_POSITION)
    m_PicturePosition = New_PicturePosition
    PropertyChanged "PicturePosition"
    Refresh
End Property
Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    Set m_PictureOriginal = New_Picture
    PropertyChanged "Picture"
    If m_PictureSize = sizeDefault Then
        m_PictureWidth = UserControl.ScaleX(m_Picture.Width, vbHimetric, UserControl.ScaleMode)
        m_PictureHeight = UserControl.ScaleY(m_Picture.Height, vbHimetric, UserControl.ScaleMode)
    End If
    Refresh
End Property

Public Property Get PictureWidth() As Long
    PictureWidth = m_PictureWidth
End Property
Public Property Let PictureWidth(ByVal New_PictureWidth As Long)
    m_PictureWidth = New_PictureWidth
    PropertyChanged "PictureWidth"
    Refresh
End Property
Public Property Get PictureHeight() As Long
    PictureHeight = m_PictureHeight
End Property
Public Property Let PictureHeight(ByVal New_PictureHeight As Long)
    m_PictureHeight = New_PictureHeight
    PropertyChanged "PictureHeight"
    Refresh
End Property
Public Property Get PictureSize() As GB_PIC_SIZE
    PictureSize = m_PictureSize
End Property
Public Property Let PictureSize(ByVal New_PictureSize As GB_PIC_SIZE)
    m_PictureSize = New_PictureSize
    PropertyChanged "PictureSize"
    Select Case New_PictureSize
    Case size16x16
        m_PictureWidth = 16
        m_PictureHeight = 16
    Case size32x32
        m_PictureWidth = 32
        m_PictureHeight = 32
    Case sizeDefault
        If Not (m_Picture Is Nothing) Then
            m_PictureWidth = UserControl.ScaleX(m_Picture.Width, vbHimetric, UserControl.ScaleMode)
            m_PictureHeight = UserControl.ScaleY(m_Picture.Height, vbHimetric, UserControl.ScaleMode)
        Else
            m_PictureWidth = 32
            m_PictureHeight = 32
        End If
    End Select
    Refresh
End Property

Private Sub CalculateCaptionRect()
    Dim mvarWidth, mvarHeight As Long
    Dim mvarFormat As Long
    With mvarDrawTextParams
        .iLeftMargin = 1
        .iRightMargin = 1
        .iTabLength = 1
        .cbSize = Len(mvarDrawTextParams)
    End With
    mvarFormat = &H400 Or &H10 Or &H4 Or &H1
    DrawTextEx UserControl.hdc, m_Caption, Len(m_Caption), mvarCaptionRect, mvarFormat, mvarDrawTextParams
    mvarWidth = mvarCaptionRect.Right - mvarCaptionRect.Left
    mvarHeight = mvarCaptionRect.Bottom - mvarCaptionRect.Top
    With mvarCaptionRect
        .Left = mvarClientRect.Left + (((mvarClientRect.Right - mvarClientRect.Left) - (mvarCaptionRect.Right - mvarCaptionRect.Left)) \ 2)
        .Top = mvarClientRect.Top + (((mvarClientRect.Bottom - mvarClientRect.Top) - (mvarCaptionRect.Bottom - mvarCaptionRect.Top)) \ 2)
        .Right = mvarCaptionRect.Left + mvarWidth
        .Bottom = mvarCaptionRect.Top + mvarHeight
    End With
End Sub

Private Sub DrawCaption()
    If m_Caption = "" Then Exit Sub
    Dim mvarForeColor As OLE_COLOR
    mvarTempRect = mvarCaptionRect
    If gb_MOUSE_IS_DOWN Then
       With mvarCaptionRect
        .Left = mvarCaptionRect.Left + 1
        .Top = mvarCaptionRect.Top + 1
        .Right = mvarCaptionRect.Right + 1
        .Bottom = mvarCaptionRect.Bottom + 1
       End With
    End If
    
    If Not Enabled Then
        Dim g_tmpFontColor As OLE_COLOR
        g_tmpFontColor = UserControl.ForeColor
        
        'AÇIK DISABLED YAZI
        UserControl.ForeColor = g_HighLight
        Dim mvarCaptionRect_Iki As RECT
        With mvarCaptionRect_Iki
            .Bottom = mvarCaptionRect.Bottom
            .Left = mvarCaptionRect.Left + 1
            .Right = mvarCaptionRect.Right + 1
            .Top = mvarCaptionRect.Top + 1
        End With
        DrawTextEx UserControl.hdc, m_Caption, Len(m_Caption), mvarCaptionRect_Iki, &H10 Or &H4 Or &H1, mvarDrawTextParams
        
        'DARK DISABLED
        UserControl.ForeColor = g_Shadow
        DrawTextEx UserControl.hdc, m_Caption, Len(m_Caption), mvarCaptionRect, &H10 Or &H4 Or &H1, mvarDrawTextParams
        
        'NORMAL
        UserControl.ForeColor = g_tmpFontColor
        Exit Sub
    End If
    
    DrawTextEx UserControl.hdc, m_Caption, Len(m_Caption), mvarCaptionRect, &H10 Or &H4 Or &H1, mvarDrawTextParams
    mvarCaptionRect = mvarTempRect
End Sub


Private Sub DrawPicture()
    Dim mvarImageType As Long
    Dim mvarImageState As Long
    Dim mvarImageFlag As Long
    If m_Picture Is Nothing Then Exit Sub
    Select Case m_Picture.Type
    Case vbPicTypeBitmap
        mvarImageType = &H4
    Case vbPicTypeIcon
        mvarImageType = &H3
    End Select
    If Not Enabled Then
        mvarImageState = &H20
    Else
        mvarImageState = &H0
    End If
    mvarTempRect = mvarPictureRect
    If gb_MOUSE_IS_DOWN Then
        With mvarPictureRect
        .Left = mvarPictureRect.Left + 1
        .Top = mvarPictureRect.Top + 1
        .Right = mvarPictureRect.Right + 1
        .Bottom = mvarPictureRect.Bottom + 1
        End With
    End If
    mvarImageFlag = mvarImageType Or mvarImageState
    PICT.Width = UserControl.ScaleX(m_Picture.Width, vbHimetric, UserControl.ScaleMode)
    PICT.Height = UserControl.ScaleY(m_Picture.Height, vbHimetric, UserControl.ScaleMode)
    PICT.ScaleMode = 3
    PICT.Cls
    DrawState PICT.hdc, 0, 0, m_Picture, 0, 0, 0, 0, 0, mvarImageFlag
    StretchBlt UserControl.hdc, mvarPictureRect.Left, mvarPictureRect.Top, mvarPictureRect.Right - mvarPictureRect.Left, mvarPictureRect.Bottom - mvarPictureRect.Top, PICT.hdc, PICT.ScaleLeft, PICT.ScaleTop, PICT.ScaleWidth, PICT.ScaleHeight, &HCC0020
    mvarPictureRect = mvarTempRect
End Sub

Public Property Get PictureHover() As Picture
    Set PictureHover = m_PictureHover
End Property

Public Property Set PictureHover(ByVal New_PictureHover As Picture)
    Set m_PictureHover = New_PictureHover
    PropertyChanged "PictureHover"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Refresh
End Property
