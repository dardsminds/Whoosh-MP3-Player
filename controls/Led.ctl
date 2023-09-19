VERSION 5.00
Begin VB.UserControl led 
   BackColor       =   &H00939393&
   BackStyle       =   0  'Transparent
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LockControls    =   -1  'True
   ScaleHeight     =   315
   ScaleWidth      =   1410
   ToolboxBitmap   =   "Led.ctx":0000
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   195
      Left            =   660
      TabIndex        =   0
      Top             =   30
      Width           =   555
   End
   Begin VB.Image imgOff 
      Height          =   195
      Left            =   0
      Picture         =   "Led.ctx":0312
      Top             =   0
      Width           =   495
   End
   Begin VB.Image imgON 
      Height          =   195
      Left            =   0
      Picture         =   "Led.ctx":0868
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "led"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Just a small button I made :)

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private mpoiCursorPos As POINTAPI

Public Event Click()
Public Event ButtonDown()
Public Event ButtonUp()

Private LedStat As Boolean

Public Event MouseOver()

Private Sub imgOff_Click()
    LedStat = True
    Status = LedStat
    UserControl.PropertyChanged "Status"
    RaiseEvent Click
End Sub


Private Sub imgON_Click()
    LedStat = False
    Status = LedStat
    UserControl.PropertyChanged "Status"
    RaiseEvent Click
End Sub



Private Sub UserControl_InitProperties()
    Status = False
    Caption = Ambient.DisplayName
End Sub

Private Sub UserControl_Resize()
    lblCaption.Left = imgOff.Width + 100
End Sub

Public Property Get Status() As Boolean
    Status = LedStat
End Property

Public Property Let Status(ByVal sNewValue As Boolean)
    LedStat = sNewValue
    UserControl.PropertyChanged "Status"
    If LedStat = True Then
        imgON.visible = True
        imgOff.visible = False
    Else
        imgON.visible = False
        imgOff.visible = True
    End If
End Property


Public Property Get Caption() As String
Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal sNewValue As String)
lblCaption.Caption = sNewValue
UserControl.PropertyChanged "Caption"
End Property

Public Property Get ForeColor() As Long
ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal lNewValue As Long)
lblCaption.ForeColor = lNewValue
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get Bold() As Boolean
Bold = lblCaption.FontBold
End Property

Public Property Let Bold(ByVal bNewValue As Boolean)
lblCaption.FontBold = bNewValue
UserControl.PropertyChanged "Bold"
End Property



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Status", LedStat, False
    .WriteProperty "Caption", Caption, Ambient.DisplayName
    .WriteProperty "ForeColor", ForeColor, 0
    .WriteProperty "Bold", Bold, False
End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    Status = .ReadProperty("Status", False)
    Caption = .ReadProperty("Caption", Ambient.DisplayName)
    ForeColor = .ReadProperty("ForeColor", 0)
    Bold = .ReadProperty("Bold", False)
End With
End Sub
