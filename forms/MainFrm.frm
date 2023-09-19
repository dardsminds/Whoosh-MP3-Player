VERSION 5.00
Object = "{34F681D0-3640-11CF-9294-00AA00B8A733}#1.0#0"; "danim.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainFrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8700
   ClientLeft      =   15
   ClientTop       =   300
   ClientWidth     =   11970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00484815&
   Icon            =   "MainFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MainFrm.frx":0442
   ScaleHeight     =   8700
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrVol 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6870
      Top             =   4920
   End
   Begin Whoosh.bigButton bigButton1 
      Height          =   570
      Left            =   5640
      TabIndex        =   119
      Top             =   3690
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1005
   End
   Begin MSComctlLib.ListView ListView1 
      DragIcon        =   "MainFrm.frx":E880
      Height          =   4140
      Left            =   30
      TabIndex        =   72
      Top             =   4500
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   7303
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   8421504
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ETA"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Artist"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "BPM"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ID"
         Object.Width           =   9
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Genre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Rating"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Last Play"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "len"
         Object.Width           =   882
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10260
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":ECC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":EEF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":F12A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":F35E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":F592
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E49667&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   3735
      TabIndex        =   73
      Top             =   2820
      Width           =   3765
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFFFF&
         Height          =   630
         Left            =   90
         TabIndex        =   113
         Top             =   270
         Width           =   3570
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total playing time:"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   60
         TabIndex        =   75
         Top             =   30
         Width           =   1380
      End
      Begin VB.Label lblTotalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1470
         TabIndex        =   74
         Top             =   0
         Width           =   750
      End
   End
   Begin MSComctlLib.ListView ListView2 
      DragIcon        =   "MainFrm.frx":F7C6
      Height          =   3945
      Left            =   5025
      TabIndex        =   71
      Top             =   4695
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   6959
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   8421504
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Artist"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Len"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ID"
         Object.Width           =   9
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Genre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Rating"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Last Play"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "len"
         Object.Width           =   882
      EndProperty
   End
   Begin Whoosh.led led1 
      Height          =   195
      Left            =   1965
      TabIndex        =   60
      Top             =   2160
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   344
      ForeColor       =   -2147483630
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   6750
      Top             =   150
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00DA9A61&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   7830
      ScaleHeight     =   1680
      ScaleWidth      =   3810
      TabIndex        =   28
      Top             =   240
      Width           =   3810
      Begin VB.PictureBox analyzer 
         Appearance      =   0  'Flat
         BackColor       =   &H00E49667&
         BorderStyle     =   0  'None
         FillColor       =   &H00E8B479&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   990
         Index           =   1
         Left            =   15
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   174
         TabIndex        =   29
         Top             =   15
         Width           =   2610
      End
      Begin VB.Label lblRem 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         Height          =   180
         Index           =   1
         Left            =   3015
         TabIndex        =   59
         Top             =   600
         Width           =   645
      End
      Begin VB.Label lblBPM 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   2985
         TabIndex        =   46
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "BPM"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   105
         Index           =   17
         Left            =   2715
         TabIndex        =   45
         Top             =   885
         Width           =   195
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "REM"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   105
         Index           =   16
         Left            =   2715
         TabIndex        =   44
         Top             =   690
         Width           =   195
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "TRIG"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   75
         Index           =   15
         Left            =   2655
         TabIndex        =   43
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "LEN"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   14
         Left            =   2715
         TabIndex        =   42
         Top             =   60
         Width           =   225
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CUR"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   105
         Index           =   13
         Left            =   2715
         TabIndex        =   41
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "TITLE"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   12
         Left            =   30
         TabIndex        =   40
         Top             =   1065
         Width           =   345
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ARTIST"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   11
         Left            =   15
         TabIndex        =   39
         Top             =   1290
         Width           =   435
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ALBUM"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   10
         Left            =   30
         TabIndex        =   38
         Top             =   1500
         Width           =   435
      End
      Begin VB.Label lblAlbum 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   540
         TabIndex        =   37
         Top             =   1455
         Width           =   1755
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "GENRE"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   9
         Left            =   2325
         TabIndex        =   36
         Top             =   1530
         Width           =   360
      End
      Begin VB.Label lblGenre 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   2655
         TabIndex        =   35
         Top             =   1455
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   34
         Top             =   1035
         Width           =   3120
      End
      Begin VB.Label lblArtist 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   33
         Top             =   1245
         Width           =   3120
      End
      Begin VB.Label lblLen 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         Caption         =   "00:00:00"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   3000
         TabIndex        =   32
         Top             =   15
         Width           =   675
      End
      Begin VB.Label lblTrig 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         Caption         =   "00:00:00"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   3000
         TabIndex        =   31
         Top             =   210
         Width           =   675
      End
      Begin VB.Label lblCurTime 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         Caption         =   "00:00:00"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   3000
         TabIndex        =   30
         Top             =   420
         Width           =   675
      End
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   5220
      Top             =   4845
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":FC08
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":1005C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":104B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":10904
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":10D58
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":111AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":11600
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":11A54
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":11EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":122FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":12750
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":12BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":12FF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":1344C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":138A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFrm.frx":13CF4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   5670
      Top             =   120
   End
   Begin VB.Timer tmrClock 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6150
      Top             =   150
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E49667&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   300
      ScaleHeight     =   1680
      ScaleWidth      =   3690
      TabIndex        =   5
      Top             =   240
      Width           =   3690
      Begin VB.PictureBox analyzer 
         Appearance      =   0  'Flat
         BackColor       =   &H00E49667&
         BorderStyle     =   0  'None
         FillColor       =   &H00E8B479&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   990
         Index           =   0
         Left            =   30
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   174
         TabIndex        =   118
         Top             =   30
         Width           =   2610
      End
      Begin VB.Label lblRem 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         Height          =   180
         Index           =   0
         Left            =   3000
         TabIndex        =   58
         Top             =   630
         Width           =   645
      End
      Begin VB.Label lblCurTime 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         Caption         =   "00:00:00"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   3000
         TabIndex        =   27
         Top             =   420
         Width           =   675
      End
      Begin VB.Label lblTrig 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         Caption         =   "00:00:00"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   3000
         TabIndex        =   26
         Top             =   210
         Width           =   675
      End
      Begin VB.Label lblLen 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         Caption         =   "00:00:00"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   3000
         TabIndex        =   25
         Top             =   15
         Width           =   675
      End
      Begin VB.Label lblArtist 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   23
         Top             =   1245
         Width           =   3120
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   22
         Top             =   1035
         Width           =   3120
      End
      Begin VB.Label lblGenre 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   2655
         TabIndex        =   18
         Top             =   1455
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "GENRE"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   8
         Left            =   2325
         TabIndex        =   17
         Top             =   1530
         Width           =   360
      End
      Begin VB.Label lblAlbum 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   540
         TabIndex        =   16
         Top             =   1455
         Width           =   1755
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ALBUM"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   7
         Left            =   30
         TabIndex        =   15
         Top             =   1500
         Width           =   435
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ARTIST"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   1
         Left            =   15
         TabIndex        =   13
         Top             =   1290
         Width           =   435
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "TITLE"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   0
         Left            =   30
         TabIndex        =   12
         Top             =   1065
         Width           =   345
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CUR"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   105
         Index           =   5
         Left            =   2715
         TabIndex        =   11
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "LEN"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   135
         Index           =   4
         Left            =   2715
         TabIndex        =   10
         Top             =   60
         Width           =   225
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "TRIG"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   75
         Index           =   3
         Left            =   2655
         TabIndex        =   9
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "REM"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   105
         Index           =   2
         Left            =   2715
         TabIndex        =   8
         Top             =   690
         Width           =   195
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "BPM"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   105
         Index           =   6
         Left            =   2715
         TabIndex        =   7
         Top             =   885
         Width           =   195
      End
      Begin VB.Label lblBPM 
         Alignment       =   2  'Center
         BackColor       =   &H00E49667&
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   2985
         TabIndex        =   6
         Top             =   840
         Width           =   675
      End
   End
   Begin VB.PictureBox gph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      DrawStyle       =   1  'Dash
      FillStyle       =   2  'Horizontal Line
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1425
      Left            =   4230
      ScaleHeight     =   1395
      ScaleWidth      =   180
      TabIndex        =   4
      Top             =   4560
      Width           =   210
   End
   Begin Whoosh.cpvSlider pitchSlider 
      Height          =   1560
      Left            =   4500
      Top             =   210
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   423
      BackColor       =   14980711
      SliderIcon      =   "MainFrm.frx":14148
      RailPicture     =   "MainFrm.frx":143DA
      RailStyle       =   99
      Value           =   5
   End
   Begin Whoosh.cpvSlider pos 
      Height          =   120
      Left            =   300
      Top             =   1920
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   212
      BackColor       =   14980711
      SliderIcon      =   "MainFrm.frx":143F6
      Orientation     =   0
      RailPicture     =   "MainFrm.frx":148C8
      RailStyle       =   99
      ShowValueTip    =   0   'False
      Max             =   100
   End
   Begin Whoosh.GurhanButton StopDeck 
      Height          =   345
      Left            =   1215
      TabIndex        =   1
      ToolTipText     =   "Stop current playing track"
      Top             =   2160
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      Caption         =   ""
      Picture         =   "MainFrm.frx":148E4
      PictureWidth    =   29
      PictureHeight   =   23
      PictureSize     =   2
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Whoosh.GurhanButton DeckPlay 
      Height          =   345
      Left            =   750
      TabIndex        =   2
      ToolTipText     =   "Play loaded track"
      Top             =   2160
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      Caption         =   ""
      Picture         =   "MainFrm.frx":1511E
      PictureWidth    =   29
      PictureHeight   =   23
      PictureSize     =   2
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Whoosh.GurhanButton DeckLoad 
      Height          =   345
      Left            =   315
      TabIndex        =   3
      ToolTipText     =   "Load next track from playlist"
      Top             =   2160
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      Caption         =   ""
      Picture         =   "MainFrm.frx":15958
      PictureWidth    =   29
      PictureHeight   =   23
      PictureSize     =   2
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Whoosh.cpvSlider CrossFader 
      Height          =   210
      Left            =   5175
      Top             =   2355
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   370
      BackColor       =   9079434
      SliderIcon      =   "MainFrm.frx":16192
      Orientation     =   0
      RailPicture     =   "MainFrm.frx":1652C
      RailStyle       =   99
      Max             =   100
      Value           =   50
   End
   Begin Whoosh.cpvSlider pitchb 
      Height          =   1500
      Left            =   7500
      Top             =   270
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   423
      BackColor       =   14980711
      SliderIcon      =   "MainFrm.frx":16548
      RailPicture     =   "MainFrm.frx":167DA
      RailStyle       =   99
      Value           =   5
   End
   Begin Whoosh.cpvSlider posb 
      Height          =   120
      Left            =   7470
      Top             =   1920
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   212
      BackColor       =   14980711
      SliderIcon      =   "MainFrm.frx":167F6
      Orientation     =   0
      RailPicture     =   "MainFrm.frx":16CC8
      RailStyle       =   99
      ShowValueTip    =   0   'False
      Max             =   100
   End
   Begin Whoosh.GurhanButton cmdStop 
      Height          =   345
      Left            =   8385
      TabIndex        =   19
      ToolTipText     =   "Stop current playing track"
      Top             =   2160
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      Caption         =   ""
      Picture         =   "MainFrm.frx":16CE4
      PictureWidth    =   29
      PictureHeight   =   23
      PictureSize     =   2
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Whoosh.GurhanButton cmdPlayb 
      Height          =   345
      Left            =   7935
      TabIndex        =   20
      ToolTipText     =   "Play loaded track"
      Top             =   2160
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      Caption         =   ""
      Picture         =   "MainFrm.frx":1751E
      PictureWidth    =   29
      PictureHeight   =   23
      PictureSize     =   2
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Whoosh.GurhanButton cmdLoadB 
      Height          =   345
      Left            =   7500
      TabIndex        =   21
      ToolTipText     =   "Load next track from playlist"
      Top             =   2160
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      Caption         =   ""
      Picture         =   "MainFrm.frx":17D58
      PictureWidth    =   29
      PictureHeight   =   23
      PictureSize     =   2
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Whoosh.led led2 
      Height          =   195
      Left            =   2520
      TabIndex        =   61
      Top             =   2160
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   344
      ForeColor       =   -2147483630
   End
   Begin Whoosh.led led3 
      Height          =   195
      Left            =   3075
      TabIndex        =   62
      Top             =   2160
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   344
      ForeColor       =   -2147483630
   End
   Begin Whoosh.led led4 
      Height          =   195
      Left            =   3630
      TabIndex        =   63
      Top             =   2160
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   344
      ForeColor       =   -2147483630
   End
   Begin Whoosh.led led5 
      Height          =   195
      Left            =   9060
      TabIndex        =   64
      Top             =   2160
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   344
      ForeColor       =   -2147483630
   End
   Begin Whoosh.led led6 
      Height          =   195
      Left            =   9600
      TabIndex        =   65
      Top             =   2160
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   344
      ForeColor       =   -2147483630
   End
   Begin Whoosh.led led7 
      Height          =   195
      Left            =   10155
      TabIndex        =   66
      Top             =   2160
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   344
      ForeColor       =   -2147483630
   End
   Begin Whoosh.led led8 
      Height          =   195
      Left            =   10710
      TabIndex        =   67
      Top             =   2160
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   344
      ForeColor       =   -2147483630
   End
   Begin Whoosh.led led9 
      Height          =   195
      Left            =   4215
      TabIndex        =   68
      Top             =   2160
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   344
      ForeColor       =   -2147483630
   End
   Begin Whoosh.led led10 
      Height          =   195
      Left            =   11310
      TabIndex        =   69
      Top             =   2160
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   344
      ForeColor       =   -2147483630
   End
   Begin MSComctlLib.ListView ListView3 
      DragIcon        =   "MainFrm.frx":18592
      Height          =   4155
      Left            =   9960
      TabIndex        =   70
      Top             =   4485
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   8421504
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Category"
         Object.Width           =   2999
      EndProperty
   End
   Begin Whoosh.dbutton dbutton1 
      Height          =   225
      Left            =   5760
      TabIndex        =   76
      ToolTipText     =   "Fade to next song"
      Top             =   2760
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   370
   End
   Begin Whoosh.led Browser 
      Height          =   225
      Left            =   3390
      TabIndex        =   77
      Top             =   2580
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   397
      Status          =   -1  'True
      Caption         =   " "
      ForeColor       =   8388608
   End
   Begin Whoosh.led eqEnable 
      Height          =   210
      Left            =   11190
      TabIndex        =   78
      Top             =   3060
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   370
      Caption         =   "led11"
      ForeColor       =   -2147483630
   End
   Begin Whoosh.Equalizer Equalizer1 
      Height          =   1125
      Index           =   0
      Left            =   9960
      TabIndex        =   79
      Top             =   3270
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1984
      Frequency       =   80
      Value           =   7
      LedOnColor      =   65280
   End
   Begin Whoosh.Equalizer Equalizer1 
      Height          =   1125
      Index           =   1
      Left            =   10215
      TabIndex        =   80
      Top             =   3270
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1984
      Frequency       =   250
      Value           =   7
      LedOnColor      =   65280
   End
   Begin Whoosh.Equalizer Equalizer1 
      Height          =   1125
      Index           =   2
      Left            =   10470
      TabIndex        =   81
      Top             =   3270
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1984
      Frequency       =   500
      Value           =   7
      LedOnColor      =   65280
   End
   Begin Whoosh.Equalizer Equalizer1 
      Height          =   1125
      Index           =   3
      Left            =   10725
      TabIndex        =   82
      Top             =   3270
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1984
      Value           =   7
      LedOnColor      =   65280
   End
   Begin Whoosh.Equalizer Equalizer1 
      Height          =   1125
      Index           =   4
      Left            =   10980
      TabIndex        =   83
      Top             =   3270
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1984
      Frequency       =   2000
      Value           =   7
      LedOnColor      =   65280
   End
   Begin Whoosh.Equalizer Equalizer1 
      Height          =   1125
      Index           =   5
      Left            =   11235
      TabIndex        =   84
      Top             =   3270
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1984
      Frequency       =   5000
      Value           =   7
      LedOnColor      =   65280
   End
   Begin Whoosh.Equalizer Equalizer1 
      Height          =   1125
      Index           =   6
      Left            =   11475
      TabIndex        =   85
      Top             =   3270
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1984
      Frequency       =   8000
      Value           =   7
      LedOnColor      =   65280
   End
   Begin Whoosh.led loops 
      Height          =   225
      Index           =   0
      Left            =   7770
      TabIndex        =   86
      Top             =   2820
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   397
      Caption         =   "led11"
   End
   Begin Whoosh.led loops 
      Height          =   225
      Index           =   2
      Left            =   8505
      TabIndex        =   87
      Top             =   2820
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   397
      Caption         =   "led11"
   End
   Begin Whoosh.cpvSlider loopspeed 
      Height          =   120
      Left            =   7800
      Top             =   3465
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   212
      BackColor       =   9079434
      SliderIcon      =   "MainFrm.frx":189D4
      Orientation     =   0
      RailPicture     =   "MainFrm.frx":18EA6
      RailStyle       =   99
      ShowValueTip    =   0   'False
      Max             =   100
   End
   Begin Whoosh.led loops 
      Height          =   225
      Index           =   1
      Left            =   7770
      TabIndex        =   88
      Top             =   3120
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   397
      Caption         =   "led11"
   End
   Begin Whoosh.led loops 
      Height          =   225
      Index           =   3
      Left            =   8505
      TabIndex        =   89
      Top             =   3120
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   397
      Caption         =   "led11"
   End
   Begin Whoosh.dbutton dbutton2 
      Height          =   225
      Left            =   210
      TabIndex        =   98
      ToolTipText     =   "Playlist Generator"
      Top             =   4020
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   370
   End
   Begin Whoosh.dbutton dbutton3 
      Height          =   225
      Left            =   9210
      TabIndex        =   99
      ToolTipText     =   "MP3 browser"
      Top             =   3510
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   370
   End
   Begin Whoosh.cpvSlider volMaster 
      Height          =   1410
      Left            =   7350
      Top             =   2820
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   423
      BackColor       =   12632256
      SliderIcon      =   "MainFrm.frx":18EC2
      RailPicture     =   "MainFrm.frx":191EE
      RailStyle       =   99
      Max             =   100
   End
   Begin Whoosh.dbutton dbutton4 
      Height          =   225
      Left            =   5040
      TabIndex        =   105
      ToolTipText     =   "Play preview"
      Top             =   4470
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   370
      MODES           =   1
   End
   Begin Whoosh.dbutton dbutton5 
      Height          =   225
      Left            =   5670
      TabIndex        =   106
      ToolTipText     =   "Stop preview"
      Top             =   4470
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   370
      MODES           =   2
   End
   Begin Whoosh.dbutton samp 
      Height          =   225
      Index           =   0
      Left            =   4170
      TabIndex        =   108
      ToolTipText     =   "Playlist Generator"
      Top             =   2820
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   370
   End
   Begin Whoosh.dbutton samp 
      Height          =   225
      Index           =   1
      Left            =   4170
      TabIndex        =   109
      ToolTipText     =   "Playlist Generator"
      Top             =   3630
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   370
   End
   Begin Whoosh.dbutton samp 
      Height          =   225
      Index           =   2
      Left            =   4170
      TabIndex        =   110
      ToolTipText     =   "Playlist Generator"
      Top             =   3210
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   370
   End
   Begin Whoosh.dbutton samp 
      Height          =   225
      Index           =   3
      Left            =   4170
      TabIndex        =   111
      ToolTipText     =   "Playlist Generator"
      Top             =   4020
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   370
   End
   Begin Whoosh.dbutton dbutton6 
      Height          =   225
      Left            =   9210
      TabIndex        =   114
      ToolTipText     =   "Show mixpoint editor"
      Top             =   4110
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   370
   End
   Begin VB.PictureBox abuff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00DA9A61&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   1380
      ScaleHeight     =   1875
      ScaleWidth      =   2685
      TabIndex        =   14
      Top             =   4590
      Width           =   2715
   End
   Begin Whoosh.led ledSmoth 
      Height          =   225
      Left            =   6600
      TabIndex        =   123
      Top             =   4020
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   397
      Status          =   -1  'True
      Caption         =   " "
      ForeColor       =   8388608
   End
   Begin Whoosh.dbutton cmdSave 
      Height          =   225
      Left            =   2580
      TabIndex        =   125
      ToolTipText     =   "Playlist Generator"
      Top             =   4020
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   370
   End
   Begin Whoosh.dbutton cmdLoad 
      Height          =   225
      Left            =   3270
      TabIndex        =   127
      ToolTipText     =   "Playlist Generator"
      Top             =   4020
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   370
   End
   Begin Whoosh.dbutton cmdClear 
      Height          =   225
      Left            =   1890
      TabIndex        =   129
      ToolTipText     =   "Playlist Generator"
      Top             =   4020
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   370
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2010
      TabIndex        =   130
      Top             =   3840
      Width           =   525
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "LOAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3390
      TabIndex        =   128
      Top             =   3840
      Width           =   465
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2700
      TabIndex        =   126
      Top             =   3840
      Width           =   525
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "SLIDE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6600
      TabIndex        =   124
      Top             =   3840
      Width           =   465
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "FX4"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   4350
      TabIndex        =   122
      Top             =   3900
      Width           =   330
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "FX3"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   4350
      TabIndex        =   121
      Top             =   3480
      Width           =   330
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "FX2"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   4350
      TabIndex        =   120
      Top             =   3090
      Width           =   330
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "PITCH"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   7170
      TabIndex        =   117
      Top             =   960
      Width           =   390
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "PITCH"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   4710
      TabIndex        =   116
      Top             =   930
      Width           =   345
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "MIXPOINT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9210
      TabIndex        =   115
      Top             =   3900
      Width           =   660
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   180
      Left            =   90
      TabIndex        =   112
      Top             =   2610
      Width           =   750
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "AUTO DJ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   2610
      TabIndex        =   107
      Top             =   2580
      Width           =   690
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "PREVIEW"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   150
      Left            =   6330
      TabIndex        =   104
      Top             =   4500
      Width           =   690
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "ETA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   150
      TabIndex        =   103
      Top             =   4290
      Width           =   420
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MAIN VOL."
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   7260
      TabIndex        =   102
      Top             =   4290
      Width           =   465
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "PLAYLIST GEN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   101
      Top             =   3840
      Width           =   915
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "BROWSER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9210
      TabIndex        =   100
      Top             =   3300
      Width           =   660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "GRAPHIC EQ."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   9960
      TabIndex        =   97
      Top             =   3090
      Width           =   1020
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "BEAT LOOPS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   7800
      TabIndex        =   96
      Top             =   2640
      Width           =   1110
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "LOOP SPEED"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   75
      Left            =   8325
      TabIndex        =   95
      Top             =   3705
      Width           =   795
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "FX1"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   4350
      TabIndex        =   94
      Top             =   2670
      Width           =   330
   End
   Begin VB.Label lblTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "12:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   10440
      TabIndex        =   93
      Top             =   2610
      Width           =   1275
   End
   Begin VB.Label lblMemory 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4770
      TabIndex        =   92
      Top             =   330
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Memory:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4770
      TabIndex        =   91
      Top             =   150
      Width           =   690
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "FADE TO NEXT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   5520
      TabIndex        =   90
      Top             =   3000
      Width           =   1170
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "AGC"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   11505
      TabIndex        =   57
      Top             =   2385
      Width           =   240
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "AGC"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   4380
      TabIndex        =   56
      Top             =   2385
      Width           =   240
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "REVERB"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   9135
      TabIndex        =   55
      Top             =   2385
      Width           =   390
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "FLANGER"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   9630
      TabIndex        =   54
      Top             =   2385
      Width           =   450
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "ECHO"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   10275
      TabIndex        =   53
      Top             =   2385
      Width           =   450
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "HI-PITCH"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   10695
      TabIndex        =   52
      Top             =   2385
      Width           =   510
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "HI-PITCH"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   3615
      TabIndex        =   51
      Top             =   2385
      Width           =   510
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "ECHO"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   3180
      TabIndex        =   50
      Top             =   2385
      Width           =   450
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "FLANGER"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   2565
      TabIndex        =   49
      Top             =   2385
      Width           =   450
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "REVERB"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   2040
      TabIndex        =   48
      Top             =   2385
      Width           =   390
   End
   Begin VB.Label lblPlayer2title 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "DECK B : IDLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A25617&
      Height          =   180
      Left            =   8475
      TabIndex        =   47
      Top             =   90
      Width           =   1860
   End
   Begin DirectAnimationCtl.DAViewerControl DAControl 
      Height          =   1440
      Left            =   5370
      TabIndex        =   24
      Top             =   660
      Width           =   1440
      OpaqueForHitDetect=   -1  'True
      UpdateInterval  =   0.033
   End
   Begin VB.Label lblPlayer1title 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "DECK A : IDLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A25617&
      Height          =   180
      Left            =   1290
      TabIndex        =   0
      Top             =   90
      Width           =   1860
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuplgenerator 
         Caption         =   "Playlist generator"
      End
      Begin VB.Menu mnubrowser 
         Caption         =   "Mp3 browser"
      End
      Begin VB.Menu mnumixpointedit 
         Caption         =   "Mixpoint Editor"
      End
   End
   Begin VB.Menu mnuoption 
      Caption         =   "Options"
      Begin VB.Menu mnuaudioconfig 
         Caption         =   "Audio configuration"
      End
      Begin VB.Menu mnusqlservice 
         Caption         =   "Database SQL service"
      End
      Begin VB.Menu sepoption 
         Caption         =   "-"
      End
      Begin VB.Menu mnumovetohistory 
         Caption         =   "Move played music to history"
      End
      Begin VB.Menu mnuautoplay 
         Caption         =   "Auto-Play at startup"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "About"
      Begin VB.Menu mnuhelp 
         Caption         =   "Help"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuauthor 
         Caption         =   "Author"
      End
   End
   Begin VB.Menu workAreaMenu 
      Caption         =   "WorkAreaMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuEditMixpoint 
         Caption         =   "Edit mixpoint"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mCopyAllToPlaylist 
         Caption         =   "Copy all to Playlist"
      End
      Begin VB.Menu mReplacePlaylist 
         Caption         =   "Replace all to Playlist"
      End
      Begin VB.Menu mnuSetRating 
         Caption         =   "Set Rating as "
         Begin VB.Menu mnu1 
            Caption         =   "1"
         End
         Begin VB.Menu mnu2 
            Caption         =   "2"
         End
         Begin VB.Menu mnu3 
            Caption         =   "3"
         End
         Begin VB.Menu mnu4 
            Caption         =   "4"
         End
         Begin VB.Menu mnu5 
            Caption         =   "5"
         End
         Begin VB.Menu mnu6 
            Caption         =   "6"
         End
         Begin VB.Menu mnu7 
            Caption         =   "7"
         End
         Begin VB.Menu mnu8 
            Caption         =   "8"
         End
         Begin VB.Menu mnu9 
            Caption         =   "9"
         End
         Begin VB.Menu mnu10 
            Caption         =   "10"
         End
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Play preview  [~]"
      End
      Begin VB.Menu mnuStopPrev 
         Caption         =   "Stop preview [Escape]"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mCut 
         Caption         =   "Cut [Del]"
      End
   End
   Begin VB.Menu mnuminimize 
      Caption         =   "Minimize"
   End
   Begin VB.Menu mnufolder 
      Caption         =   "Folder"
      Visible         =   0   'False
      Begin VB.Menu mnunewfolder 
         Caption         =   "New folder"
      End
      Begin VB.Menu mnurefreshfolder 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuremovefolder 
         Caption         =   "Remove folder"
      End
      Begin VB.Menu mnurenamefolder 
         Caption         =   "Rename folder"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents Player1 As clsPlayer
Attribute Player1.VB_VarHelpID = -1
Public WithEvents Player2 As clsPlayer
Attribute Player2.VB_VarHelpID = -1

Dim rsMusic As ADODB.Recordset
Dim rsPlaylist As ADODB.Recordset
Dim rsMusicDrag As ADODB.Recordset
Dim rshistory As ADODB.Recordset
Dim rsloadhistory As ADODB.Recordset
Dim rstmpmusic As ADODB.Recordset
Dim rspreview As ADODB.Recordset

Dim BgMusic As MUSICFILE
Dim SourceRow As Integer
Dim memIndex As Integer
Dim dest As Integer
Dim previewchan As Long
Dim curTime As Date
Dim tSeconds As Single
Dim sTime As clsBassTime
Dim origFolderName As String
Dim newFolder As Boolean

Public CurrentPlayer As Integer
Public idxmixpoint As Long
Dim MFade(100) As Integer
Dim MFadeB(100) As Integer

Dim LoopChan As Long
Dim BgChan As Long

Dim pfile As String
Const maxp = 1.2
Const minp = 0.8
Dim freq As Long

Dim LoadPoint As Long
Dim loopStart As Long

Dim i As Integer
Dim j As Integer
Dim file As String

Public MeReadyToStop As Boolean

Dim Player1Tmr As Integer
Dim Player2Tmr As Integer
Dim TalkActive As Boolean
Dim OrigVol As Integer
Dim TalkVol As Integer

Public Function LoadMP3_1()
Dim file As String
Dim length As Long
Dim idx As Long

    'nothing to load then exit
    If ListView1.ListItems.count < 1 Then
        ReloadHistory
        Exit Function
    End If
    
    If Timer1.Enabled = True Then
        Timer1.Enabled = False
    End If
    
    idx = ListView1.ListItems.Item(1).SubItems(4)
    ListView1.ListItems.Remove (1)
    
    Set rsMusic = Nothing
    Set rsMusic = OpenRS("SELECT * FROM music WHERE index=" & idx)
    If rsMusic.RecordCount = 0 Then Exit Function
    
    ExecuteQuery "DELETE FROM playlist WHERE MusicIndex=" & idx
    
    'save time and date last play ReloadHistory
    rsMusic.Fields!LastPlay = Now
    rsMusic.Update
    
    'save loaded song to history list
    SaveSongToHistory idx
    
    Player1.UnloadMP3
    
    With Player1
        .pfile = rsMusic.Fields!file
        .pTitle = rsMusic.Fields!Title
        .pArtist = rsMusic.Fields!Artist
        .pAlbum = rsMusic.Fields!Album
        .pGenre = rsMusic.Fields!Genre
        .pBpm = rsMusic.Fields!bpm
        .pSongStart = rsMusic.Fields!SongStart
        .pSongEnd = rsMusic.Fields!SongEnd
        .pLen = rsMusic.Fields!Len
        .pCategory = Trim(rsMusic.Fields!Category)
        
    End With
    
    If Player1.pCategory = "~VoiceOverBackground" Then
        Player1.Mode = pBACKGROUND
    Else
        Player1.Mode = pNORMAL
    End If
    
       'set pitch
    With pitchSlider
        .max = 1
        .Min = 0
        .max = (maxp - 1) * 100
        .Min = (minp - 1) * 100
        .value = 0
    End With
    
    Player1.Load
    
    
    'e preparar ang parametric ekwalayser
    SetEqualizer
    
    lblLen(0).Caption = Player1.getLengthTime()
    lblTrig(0).Caption = Player1.getMixpointTime()
    lblBPM(0).Caption = Player1.pBpm
    
    Player1.Status = STOPPED
    Player1.setPosition Player1.pSongStart
    'Player1.Play
    'Player1.PauseMp3
    
    'lblBPM.Caption = Mp3(1).orgBPM
    'MeReadyToStop = False
    'lblBPM.Caption = DecodeBPM1(True, 0, 30, Mp3(1).file.file)
    'Call BASS_FX_BPM_CallbackSet(Mp3(1).chan, AddressOf GetBPM_Callback1, 5, 0, BASS_FX_BPM_MULT2)

     
End Function

Public Function Load_MP3_1(mindex As Double) 'player1.unloadmp3
Dim file As String
Dim length As Long
Dim idx As Long
    
    If Timer1.Enabled = True Then
        Timer1.Enabled = False
    End If
    
    idx = mindex
    
    ExecuteQuery "DELETE FROM playlist WHERE MusicIndex=" & idx
    
    Set rsMusic = Nothing
    Set rsMusic = OpenRS("SELECT * FROM music WHERE index=" & idx)
    If rsMusic.RecordCount = 0 Then Exit Function
    'save time and date last play
    rsMusic.Fields!LastPlay = Now
    rsMusic.Update
    
    'save loaded song to history list
    SaveSongToHistory idx
    
    Player1.UnloadMP3
    
    With Player1
        .pfile = rsMusic.Fields!file
        .pTitle = rsMusic.Fields!Title
        .pArtist = rsMusic.Fields!Artist
        .pAlbum = rsMusic.Fields!Album
        .pGenre = rsMusic.Fields!Genre
        .pBpm = rsMusic.Fields!bpm
        .pSongStart = rsMusic.Fields!SongStart
        .pSongEnd = rsMusic.Fields!SongEnd
        .pLen = rsMusic.Fields!Len
        .pCategory = Trim(rsMusic.Fields!Category)
    End With
    
    If Player1.pCategory = "~VoiceOverBackground" Then
        Player1.Mode = pBACKGROUND
    Else
        Player1.Mode = pNORMAL
    End If
    
       'set pitch
    With pitchSlider
        .max = 1
        .Min = 0
        .max = (maxp - 1) * 100
        .Min = (minp - 1) * 100
        .value = 0
    End With
    
    Player1.Load
    SetEqualizer
    
    lblLen(0).Caption = Player1.getLengthTime()
    lblTrig(0).Caption = Player1.getMixpointTime()
    lblBPM(0).Caption = Player1.pBpm
    
    Player1.Status = STOPPED
    Player1.setPosition Player1.pSongStart
    'Player1.Play
    'Player1.PauseMp3
    
    'lblBPM.Caption = Mp3(1).orgBPM
    'MeReadyToStop = False
    'lblBPM.Caption = DecodeBPM1(True, 0, 30, Mp3(1).file.file)
    'Call BASS_FX_BPM_CallbackSet(Mp3(1).chan, AddressOf GetBPM_Callback1, 5, 0, BASS_FX_BPM_MULT2)

     
End Function

Public Function Load_MP3_2(mindex As Double)
Dim file As String
Dim length As Long
Dim idx As Long
    If Timer2.Enabled = True Then
        Timer2.Enabled = False
    End If
    idx = mindex
    ExecuteQuery "DELETE FROM playlist WHERE MusicIndex=" & idx
    Set rsMusic = Nothing
    Set rsMusic = OpenRS("SELECT * FROM music WHERE index=" & idx)
    If rsMusic.RecordCount = 0 Then Exit Function
    
    'save time and date last play
    rsMusic.Fields!LastPlay = Now
    rsMusic.Update
    
    'save loaded song to history list
    SaveSongToHistory idx
    Player2.UnloadMP3
    With Player2
        .pfile = rsMusic.Fields!file
        .pTitle = rsMusic.Fields!Title
        .pArtist = rsMusic.Fields!Artist
        .pAlbum = rsMusic.Fields!Album
        .pGenre = rsMusic.Fields!Genre
        .pBpm = rsMusic.Fields!bpm
        .pSongStart = rsMusic.Fields!SongStart
        .pSongEnd = rsMusic.Fields!SongEnd
        .pLen = rsMusic.Fields!Len
        .pCategory = Trim(rsMusic.Fields!Category)
    End With
    
    If Player2.pCategory = "~VoiceOverBackground" Then
        Player2.Mode = pBACKGROUND
    Else
        Player2.Mode = pNORMAL
    End If
    
       'set pitch
    With pitchb
        .max = 1
        .Min = 0
        .max = (maxp - 1) * 100
        .Min = (minp - 1) * 100
        .value = 0
    End With
    
    
    Player2.Load
    lblLen(1).Caption = Player2.getLengthTime()
    lblTrig(1).Caption = Player2.getMixpointTime()
    lblBPM(1).Caption = Player2.pBpm
    
    'e preparar ang parametric ekwalayser
    SetEqualizer

    Player2.Status = STOPPED
    Player2.setPosition Player2.pSongStart
    'Player2.Play
    'Player2.PauseMp3

    'lblBPM.Caption = Mp3(1).orgBPM
    'lblBPM.Caption = DecodeBPM1(True, 0, 30, Mp3(1).file.file)
    'Call BASS_FX_BPM_CallbackSet(Mp3(1).chan, AddressOf GetBPM_Callback1, 5, 0, BASS_FX_BPM_MULT2)
    
     analyzer(1).Cls
     
End Function


Public Function LoadMP3_2()
Dim file As String
Dim length As Long
Dim idx As Long

    'nothing to load then exit
    If ListView1.ListItems.count < 1 Then
        ReloadHistory
        'Exit Function
    End If
    
    
    If Timer2.Enabled = True Then
        Timer2.Enabled = False
    End If
    
    idx = ListView1.ListItems.Item(1).SubItems(4)
    ListView1.ListItems.Remove (1)
    
    ExecuteQuery "DELETE FROM playlist WHERE MusicIndex=" & idx
    
    Set rsMusic = Nothing
    Set rsMusic = OpenRS("SELECT * FROM music WHERE index=" & idx)
    If rsMusic.RecordCount = 0 Then Exit Function
    
    
    'save time and date last play
    rsMusic.Fields!LastPlay = Now
    rsMusic.Update
    
    'save loaded song to history list
    SaveSongToHistory idx
    
    Player2.UnloadMP3

    With Player2
        .pfile = rsMusic.Fields!file
        .pTitle = rsMusic.Fields!Title
        .pArtist = rsMusic.Fields!Artist
        .pAlbum = rsMusic.Fields!Album
        .pGenre = rsMusic.Fields!Genre
        .pBpm = rsMusic.Fields!bpm
        .pSongStart = rsMusic.Fields!SongStart
        .pSongEnd = rsMusic.Fields!SongEnd
        .pLen = rsMusic.Fields!Len
        .pCategory = Trim(rsMusic.Fields!Category)
    End With
    
    If Player2.pCategory = "~VoiceOverBackground" Then
        Player2.Mode = pBACKGROUND
    Else
        Player2.Mode = pNORMAL
    End If
    
       'set pitch
    With pitchb
        .max = 1
        .Min = 0
        .max = (maxp - 1) * 100
        .Min = (minp - 1) * 100
        .value = 0
    End With
    
    
    Player2.Load
    lblLen(1).Caption = Player2.getLengthTime()
    lblTrig(1).Caption = Player2.getMixpointTime()
    lblBPM(1).Caption = Player2.pBpm
    
    'e preparar ang parametric ekwalayser
    SetEqualizer

    Player2.Status = STOPPED
    Player2.setPosition Player2.pSongStart
    'Player2.Play
    'Player2.PauseMp3

    'lblBPM.Caption = Mp3(1).orgBPM
    'lblBPM.Caption = DecodeBPM1(True, 0, 30, Mp3(1).file.file)
    'Call BASS_FX_BPM_CallbackSet(Mp3(1).chan, AddressOf GetBPM_Callback1, 5, 0, BASS_FX_BPM_MULT2)
    
     analyzer(1).Cls
     
     
End Function


Private Sub SaveSongToHistory(idx As Long)
    Dim sqltxt As String
    Set rshistory = OpenRS("SELECT * FROM music WHERE index=" & idx)
    If rshistory.RecordCount <> 0 Then
        sqltxt = "INSERT INTO history(musicindex,len,SongStart,SongEnd,MixType) "
        sqltxt = sqltxt & " VALUES(" & idx & "," & rshistory.Fields!Len & "," & rshistory.Fields!SongStart & "," & rshistory.Fields!SongEnd & ",'" & rshistory.Fields!MixType & "')"
        ExecuteQuery sqltxt
    End If
End Sub

Private Sub SaveCompose(idx As Long, composename As String, djname As String)
    Dim sqltxt As String
    Set rshistory = OpenRS("SELECT * FROM music WHERE index=" & idx)
    If rshistory.RecordCount <> 0 Then
        sqltxt = "INSERT INTO composition(musicindex,len,SongStart,SongEnd,MixType,composename,dj) "
        sqltxt = sqltxt & " VALUES(" & idx & "," & rshistory.Fields!Len & "," & rshistory.Fields!SongStart & "," & rshistory.Fields!SongEnd & ",'" & rshistory.Fields!MixType & "','" & composename & "','" & djname & "')"
        ExecuteQuery sqltxt
    End If
End Sub

Private Sub analyzer_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Dim mindex As Double
    
If Index = 0 Then
    If Source.name = "ListView2" Then
        i = ListView2.SelectedItem.Index
        mindex = ListView2.ListItems.Item(i).SubItems(3)
        If mindex = -1 Then Exit Sub
        Load_MP3_1 mindex
    End If
End If

If Index = 1 Then
    If Source.name = "ListView2" Then
        i = ListView2.SelectedItem.Index
        mindex = ListView2.ListItems.Item(i).SubItems(3)
        If mindex = -1 Then Exit Sub
        Load_MP3_2 mindex
    End If
End If



End Sub


Private Sub bigButton1_Click()
    TalkActive = Not TalkActive
    TmrVol.Enabled = True
    
    If TalkActive Then
        bigButton1.MODES = 0
        OrigVol = volMaster.value
    Else
        bigButton1.MODES = 1
    End If
End Sub

Private Sub Browser_Click()
    AutoDJ = Browser.Status
End Sub

Private Sub cmdClear_Click()
    If ListView1.ListItems.count = 0 Then Exit Sub
    
    If MsgBox("Are you sure you want to clear the current playlist?", vbYesNo, "Clear Playlist") = vbNo Then Exit Sub
    ListView1.ListItems.Clear
End Sub

Private Sub cmdLoad_Click()
    frmLoadCompose.Show 1
End Sub

Private Sub cmdLoadB_Click()
    LoadMP3_2
End Sub



Private Sub cmdPlayb_Click()
    If Player2.IsLoaded = False Then
         lblStatus.Caption = "Player 2 is empty, please load song first"
         Exit Sub
    End If
    
    Player2.Play
End Sub





Private Sub cmdsave_Click()
    frmSaveCompose.Show 1
End Sub

Private Sub cmdStop_Click()
    Player2.StopPlay
End Sub





Public Sub CrossFader_ValueChanged()
    Player1.setVolume MFade(CrossFader.value)
    Player2.setVolume MFadeB(CrossFader.value)
End Sub


Private Sub dbutton1_Click()
      Call MixNow
End Sub

Private Sub dbutton2_Click()
    PLaylistGeneratorFrm.Show
End Sub

Private Sub dbutton3_Click()
    BrowserFrm.Show
End Sub

Private Sub dbutton4_Click()
    Dim pfile As String
    Dim idx As Long
    Dim sqltxt As String

'Preview song
If ListView2.ListItems.count > 0 Then
    idx = ListView2.ListItems.Item(ListView2.SelectedItem.Index).SubItems(3)
    
    Set rspreview = OpenRS("SELECT * FROM music WHERE index=" & idx)
    If rspreview.RecordCount = 0 Then Exit Sub
    pfile = rspreview.Fields!file
    If BASS_ChannelIsActive(previewchan) = BASS_ACTIVE_PLAYING Then BASS_ChannelStop previewchan
    BASS_StreamFree previewchan
    If AUDIO.EnableCueDevice = 1 Then
        Call BASS_SetDevice(AUDIO.CueDevice)
    End If
    previewchan = BASS_StreamCreateFile(BASSFALSE, pfile, 0, 0, 0)
    Call BASS_ChannelPlay(previewchan, BASSTRUE)
    dbutton4.MODES = mPLAYING
End If

End Sub

Private Sub dbutton5_Click()
'stop preview
If ListView2.ListItems.count > 0 Then
    BASS_ChannelStop previewchan
    BASS_StreamFree previewchan
    dbutton4.MODES = mPLAY
    
End If
End Sub



Private Sub dbutton6_Click()
    mnuEditMixpoint_Click
End Sub

Private Sub DeckLoad_Click()
    LoadMP3_1
End Sub

Private Sub DeckPlay_Click()
    If Player1.IsLoaded = False Then
         lblStatus.Caption = "Player 1 is empty, please load song first"
         Exit Sub
    End If
    
    Player1.Play
End Sub

Private Sub eqEnable_Click()
    SetEqualizer
End Sub

Private Sub Equalizer1_ValueChange(Index As Integer, newValue As Integer)
    If Player1.IsLoaded = True Then
        Player1.UpdateEqualizer Index, (Equalizer1(Index).value - 7) * -1
    End If
    If Player2.IsLoaded = True Then
        Player2.UpdateEqualizer Index, (Equalizer1(Index).value - 7) * -1
    End If
End Sub

Private Sub Form_Activate()
    ListView1.Refresh
    ListView2.Refresh
    ListView3.Refresh
    
    volMaster.value = modBass.BASS_GetVolume()

End Sub

Private Sub Form_Load()
Dim q As Integer
Dim c As Integer
Dim sqlservice As String


Set sTime = New clsBassTime
Set Player1 = New clsPlayer
Set Player2 = New clsPlayer

    Player1.pUser = 1
    Player2.pUser = 2

    Dim freq As Long
    'arrays of fader value A scroll
    For a = 0 To 50
    MFade(a) = 100
    Next a
    For a = 50 To 100
    MFade(a) = 100 - b
    b = b + 2
    Next a
    
    'arrays of fader value b scroll
    For a = 0 To 50
    MFadeB(a) = a * 2
    Next a
    For a = 50 To 100
    MFadeB(a) = 100
    Next a
    CurrentPlayer = 0

'initialize a sub class
'oldProc = SetWindowLongA(Me.hwnd, GWL_WNDPROC, AddressOf WndProc)

'HotKeyActivate Me.hwnd, MOD_ALT, Asc("N")   'Mix next track
'HotKeyActivate Me.hwnd, MOD_ALT, Asc("M")       'hotkey 2
'HotKeyActivate Me.hwnd, MOD_SHIFT, Asc("M")     'hotkey 3
'HotKeyActivate Me.hwnd, MOD_CONTROL, Asc("J")   'hotkey 4


ChDrive App.Path
ChDir App.Path

For q = 0 To 200
GainAdj(q) = q - 200
Next q

CrossFade = True
AutoDJ = True

'tmrClock.Enabled = True  'activates the clock

'load equalizer settings
Equalizer1(0).value = ReadINI("EQUALIZER", "BAND1", 10)
Equalizer1(1).value = ReadINI("EQUALIZER", "BAND2", 7)
Equalizer1(2).value = ReadINI("EQUALIZER", "BAND3", 5)
Equalizer1(3).value = ReadINI("EQUALIZER", "BAND4", 5)
Equalizer1(4).value = ReadINI("EQUALIZER", "BAND5", 6)
Equalizer1(5).value = ReadINI("EQUALIZER", "BAND6", 8)
Equalizer1(6).value = ReadINI("EQUALIZER", "BAND7", 13)

Equalizer1(0).Frequency = ReadINI("EQUALIZER", "BAND1FREQ", 80)
Equalizer1(1).Frequency = ReadINI("EQUALIZER", "BAND2FREQ", 250)
Equalizer1(2).Frequency = ReadINI("EQUALIZER", "BAND3FREQ", 500)
Equalizer1(3).Frequency = ReadINI("EQUALIZER", "BAND4FREQ", 1000)
Equalizer1(4).Frequency = ReadINI("EQUALIZER", "BAND5FREQ", 2000)
Equalizer1(5).Frequency = ReadINI("EQUALIZER", "BAND6FREQ", 5000)
Equalizer1(6).Frequency = ReadINI("EQUALIZER", "BAND7FREQ", 8000)

eqEnable.Status = ReadINI("EQUALIZER", "ENABLE", True)
  
led9.Status = ReadINI("AGC", "AGCenable", True)
led10.Status = ReadINI("AGC", "AGCenable", True)
  
TalkVol = ReadINI("TalkVolume", "volume", "5")


'Me.Caption = App.ProductName & " " & App.Major & "." & App.Minor & App.Revision
Me.Show 'show the main window


tmrClock.Enabled = True

  Set m = DAControl.PixelLibrary
  Set fillImg = m.ImportImage(App.Path & "\ttable.jpg")
  Set ovalImg = m.Oval(85, 85).Fill(m.DefaultLineStyle, fillImg)
  Set rotXf = m.Rotate3RateDegrees(m.Vector3(0, 0, 1), 120).ParallelTransform2().Inverse
  Set finalImg = ovalImg.Transform(rotXf)
  DAControl.Image = finalImg
  DAControl.BackgroundImage = m.SolidColorImage(m.Silver)
  DAControl.Start
  TurnTableIsActive = True

If ReadINI("Main", "InstantPlay", "0") = 1 Then
    mnuautoplay.Checked = True
        LoadUserPlaylist
        'LoadMP3_1
        'Player1.setPosition Player1.pSongStart
        'Player1.Play
End If
    
    DisplayFolderList
    ClearHistory
    
    
    'this will show/hide sqlservice for database mentainance
    'this tool allow to open database content without having to use MS Access
    sqlservice = ReadINI("Service", "sql", "0")
    If sqlservice <> "1" Then
        mnusqlservice.visible = False
    End If
    
End Sub

Private Sub SetTreeViewAttrib(c As TreeView, ByVal attrib As Long)
    Const GWL_STYLE As Long = -16
    Dim rStyle As Long
    rStyle = GetWindowLong(c.hwnd, GWL_STYLE)
    rStyle = rStyle Or attrib
    Call SetWindowLong(c.hwnd, GWL_STYLE, rStyle)
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then DragForm MainFrm.hwnd, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrClock.Enabled = False
End Sub


Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
'delte key
    If KeyCode = 46 And ListView1.ListItems.count > 0 Then
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
    End If
End Sub

Private Sub ListView1_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
        Call LVDragDropMulti(ListView1, x, y)
End Sub

Private Sub ListView1_OLEDragOver(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, state As Integer)
    Set ListView1.DropHighlight = ListView1.HitTest(x, y)
End Sub

Private Sub ListView2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu workAreaMenu
    End If
End Sub

Private Sub ListView3_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim sqltxt As String
    
    If newFolder = True Then
        sqltxt = "INSERT INTO folder(folder) VALUES('" & NewString & "')"
        ExecuteQuery sqltxt
        newFolder = False
    Else
        sqltxt = "UPDATE folder SET folder='" & NewString & "' WHERE folder='" & origFolderName & "'"
        ExecuteQuery sqltxt
        sqltxt = "UPDATE music SET Category='" & NewString & "' WHERE Category='" & origFolderName & "'"
        ExecuteQuery sqltxt
    End If

    DisplayFolderList

End Sub


Private Sub ListView3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnufolder
    End If
End Sub



Private Sub mnuaudioconfig_Click()
ConfigFrm.Show
End Sub

Private Sub mnumixpointedit_Click()
    MixPointFrm.Show
End Sub

Private Sub mnunewfolder_Click()
    Dim lst As ListItem
    Dim listcount As Integer
    
    Set lst = ListView3.ListItems.Add
    lst.text = "New folder"
    lst.SmallIcon = 1
    ListView3.ListItems(ListView3.ListItems.count).Selected = True
    ListView3.StartLabelEdit
    newFolder = True
End Sub




Private Sub mnurefreshfolder_Click()
    DisplayFolderList
End Sub

Private Sub mnuremovefolder_Click()
    Dim selectedFolder As String
    Dim sqltxt As String
    If ListView3.ListItems.count = 0 Then Exit Sub
    
    If MsgBox("This will remove selected folder and all music on it" & vbCrLf & "Continue anyway?", vbYesNo, "Delete folder") = vbNo Then Exit Sub
    
    selectedFolder = ListView3.SelectedItem.text
    sqltxt = "DELETE FROM folder WHERE folder='" & selectedFolder & "'"
    ExecuteQuery sqltxt
    
    sqltxt = "DELETE FROM music WHERE Category='" & selectedFolder & "'"
    ExecuteQuery sqltxt
    
    DisplayFolderList
End Sub

Private Sub mnurenamefolder_Click()
    If ListView3.ListItems.count = 0 Then Exit Sub
    origFolderName = ListView3.SelectedItem.text
    ListView3.StartLabelEdit
End Sub

Private Sub mnusqlservice_Click()
    frmSqlService.Show
End Sub

Private Sub mReplacePlaylist_Click()
    Dim i As Integer
    Dim mindex As Long
    Dim lst As ListItem
    ListView1.ListItems.Clear
     For i = 1 To ListView2.ListItems.count
        mindex = ListView2.ListItems.Item(i).SubItems(3)
        CopyToPlayList mindex
    Next i
    
    UpdateTimePlay
End Sub

Private Sub Picture5_DragDrop(Source As Control, x As Single, y As Single)
'If Source.Name = "ListView2" Then
'    IsDrag = True
'    LoadMP3
'    IsDrag = False
'    Set ListView2.DropHighlight = Nothing
'End If
End Sub

Private Sub Led1_Click()
If led1.Status = True Then
    Player1.effectReverb True
Else
    Player1.effectReverb False
End If
End Sub

Private Sub Led10_Click()
If led10.Status = True Then
        Player2.fxAGC = False
        Player2.fxAGC = True
Else
        Player2.fxAGC = False
End If
End Sub

Private Sub Led2_Click()
    If led2.Status = True Then
        Player1.effectFlanger True
    Else
        Player1.effectFlanger False
    End If
End Sub

Private Sub Led3_Click()
If led3.Status = True Then
        Player1.effectEcho True
Else
        Player1.effectEcho False
End If
End Sub

Private Sub Led4_Click()
    If led4.Status = True Then
         Player1.effectHighPitch True
    Else
        Player1.effectHighPitch False
    End If
End Sub

Private Sub Led5_Click()
If led5.Status = True Then
    Player2.effectReverb True
Else
    Player2.effectReverb False
End If
End Sub

Private Sub Led6_Click()
    If led6.Status = True Then
        Player2.effectFlanger True
    Else
        Player2.effectFlanger False
    End If
End Sub

Private Sub Led7_Click()
If led7.Status = True Then
        Player2.effectEcho True
Else
        Player2.effectEcho False
End If
End Sub

Private Sub Led8_Click()
    If led8.Status = True Then
         Player2.effectHighPitch True
    Else
        Player2.effectHighPitch False
    End If
End Sub

Private Sub Led9_Click()
If led9.Status = True Then
        Player1.fxAGC = False
        Player1.fxAGC = True
Else
        Player1.fxAGC = False
End If
End Sub

Private Sub ListView1_DragDrop(Source As Control, x As Single, y As Single)
Dim LS As MUSICFILE
Dim mindex As Integer
Dim i As Integer
Dim mrating As Long


If Source.name = "ListView2" Then

    i = ListView2.SelectedItem.Index
    mindex = ListView2.ListItems.Item(i).SubItems(3)
    If mindex = -1 Then Exit Sub

    Set rsMusicDrag = OpenRS("SELECT * FROM music WHERE index=" & mindex)
    If rsMusicDrag.RecordCount = 0 Then Exit Sub

    'add the rating of the song
    mrating = rsMusicDrag.Fields!Rating
    rsMusicDrag.Fields!Rating = mrating + 1
    rsMusicDrag.Update

    On Error Resume Next
    dest = ListView1.DropHighlight.Index

    If dest = 0 Then
        Set lst = ListView1.ListItems.Add()
    Else
        Set lst = ListView1.ListItems.Add(dest)
    End If
       'Lst.SmallIcon = 1
       lst.text = DateAdd("s", rsMusicDrag.Fields!Len, t)
       lst.SubItems(1) = rsMusicDrag.Fields!Title
       lst.SubItems(2) = rsMusicDrag.Fields!Artist
       lst.SubItems(3) = rsMusicDrag.Fields!bpm
       lst.SubItems(4) = mindex
       lst.SubItems(7) = rsMusicDrag.Fields!Rating
       lst.SubItems(8) = rsMusicDrag.Fields!LastPlay
       lst.SubItems(9) = rsMusicDrag.Fields!Len
      
       lst.SmallIcon = SetIcon(rsMusicDrag.Fields!Category)

       
    dest = 0
    Set ListView1.DropHighlight = Nothing
    Exit Sub
End If
End Sub




Private Sub CopyToPlayList(idx As Long)
    Set rsMusicDrag = OpenRS("SELECT * FROM music WHERE index=" & idx)
       Set lst = ListView1.ListItems.Add()
       lst.text = DateAdd("s", rsMusicDrag.Fields!Len, t)
       lst.SubItems(1) = rsMusicDrag.Fields!Title
       lst.SubItems(2) = rsMusicDrag.Fields!Artist
       lst.SubItems(3) = rsMusicDrag.Fields!bpm
       lst.SubItems(4) = idx
       lst.SubItems(7) = rsMusicDrag.Fields!Rating
       lst.SubItems(8) = rsMusicDrag.Fields!LastPlay
       lst.SubItems(9) = rsMusicDrag.Fields!Len
       lst.SmallIcon = SetIcon(rsMusicDrag.Fields!Category)
      
End Sub



Private Sub ListView1_DragOver(Source As Control, x As Single, y As Single, state As Integer)
    Set ListView1.DropHighlight = ListView1.HitTest(x, y)
End Sub

Private Sub ListView2_DblClick()
    Dim idx As Long
    idx = ListView2.ListItems.Item(ListView2.SelectedItem.Index).SubItems(3)
    AddToPlayList idx
End Sub

Private Sub AddToPlayList(MusicIndex As Long)
    Dim mindex As Long
    Dim mLen As Long
    Dim mStart As Long
    Dim mEnd As Long
    Dim MixType As String
    Dim sqltxt As String

    Me.MousePointer = vbHourglass

    Set rsMusic = OpenRS("SELECT * FROM music WHERE index=" & MusicIndex)
    If rsMusic.RecordCount = 0 Then Exit Sub
    
    mindex = rsMusic.Fields!Index
    mLen = rsMusic.Fields!Len
    mStart = rsMusic.Fields!SongStart
    mEnd = rsMusic.Fields!SongEnd
    MixType = rsMusic.Fields!MixType
    
    sqltxt = "INSERT INTO playlist(musicindex,len,SongStart,SongEnd,MixType) "
    sqltxt = sqltxt & " VALUES(" & mindex & "," & mLen & "," & mStart & "," & mEnd & ",'" & MixType & "')"
    ExecuteQuery sqltxt
    
    Me.MousePointer = vbNormal
End Sub

Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
        DisplayListByCategory Item.text
        CurrentCategory = Item.text
End Sub



Private Sub loops_Click(Index As Integer)
        PlayLoop Index, loops(Index).Status
End Sub

Private Sub loopspeed_ValueChanged()
    Call BASS_ChannelSetAttributes(LoopChan, CLng(loopspeed.value) * 4, -1, -101)
End Sub

Private Sub mnuauthor_Click()
    AboutFrm.Show
End Sub

Private Sub mnuautoplay_Click()
    mnuautoplay.Checked = Not mnuautoplay.Checked
End Sub

Private Sub mnuEditMixpoint_Click()
    If ListView2.ListItems.count > 0 Then
    idxmixpoint = ListView2.ListItems.Item(ListView2.SelectedItem.Index).SubItems(3)
    MixPointFrm.Show
    End If
End Sub

Private Sub mnuExit_Click()
    
  SaveINI "EQUALIZER", "BAND1", Equalizer1(0).value
  SaveINI "EQUALIZER", "BAND2", Equalizer1(1).value
  SaveINI "EQUALIZER", "BAND3", Equalizer1(2).value
  SaveINI "EQUALIZER", "BAND4", Equalizer1(3).value
  SaveINI "EQUALIZER", "BAND5", Equalizer1(4).value
  SaveINI "EQUALIZER", "BAND6", Equalizer1(5).value
  SaveINI "EQUALIZER", "BAND7", Equalizer1(6).value

  SaveINI "EQUALIZER", "BAND1FREQ", Equalizer1(0).Frequency
  SaveINI "EQUALIZER", "BAND2FREQ", Equalizer1(1).Frequency
  SaveINI "EQUALIZER", "BAND3FREQ", Equalizer1(2).Frequency
  SaveINI "EQUALIZER", "BAND4FREQ", Equalizer1(3).Frequency
  SaveINI "EQUALIZER", "BAND5FREQ", Equalizer1(4).Frequency
  SaveINI "EQUALIZER", "BAND6FREQ", Equalizer1(5).Frequency
  SaveINI "EQUALIZER", "BAND7FREQ", Equalizer1(6).Frequency

'  SaveINI "EQUALIZER", "BANDWIDTH", txtBandWidth.text
  SaveINI "EQUALIZER", "ENABLE", eqEnable.Status
    
  SaveINI "AGC", "AGCenable", led9.Status
    
  If mnuautoplay.Checked = True Then
    SaveINI "Main", "InstantPlay", 1
  Else
    SaveINI "Main", "InstantPlay", 0
  End If
    
  'save TalkVolume setting
  SaveINI "TalkVolume", "volume", str(TalkVol)
  
    
  Set Player1 = Nothing
  Set Player2 = Nothing
    
    CloseProgram
    
    Call BASS_Free
End Sub

Private Sub mnuminimize_Click()
    Me.WindowState = 1
End Sub







Private Sub Picture6_DragDrop(Source As Control, x As Single, y As Single)
    If Source.name = "ListView2" Then
        Dim rsBGsound As ADODB.Recordset
        i = ListView2.SelectedItem.Index
        mindex = ListView2.ListItems.Item(i).SubItems(3)
        If mindex = -1 Then Exit Sub

        Set rsBGsound = OpenRS("SELECT * FROM music WHERE index=" & mindex)
        lblBgSound.Caption = rsBGsound.Fields!Title
        BgMusic.Title = rsBGsound.Fields!Title
        BgMusic.file = rsBGsound.Fields!file
    End If
End Sub

Private Sub pitchb_ValueChanged()
    Player2.setTempo pitchb.value

End Sub

Private Sub pitchSlider_ValueChanged()
    Player1.setTempo pitchSlider.value
    'Mp3(1).newBPM = GetNewBPM1()
End Sub


Private Sub Player1_onEndMix()
    Player2.setPosition Player2.pSongStart
    Player2.Play
End Sub

Private Sub Player1_onLoadMp3()
    lblTitle(0).Caption = Player1.pTitle
    lblArtist(0).Caption = Player1.pArtist
    lblAlbum(0).Caption = Player1.pAlbum
    lblGenre(0).Caption = Player1.pGenre
    
    pos.max = Player1.getStreamlenght()
    pos.Min = 0
    lblPlayer1title.Caption = "PLAYER A: IDLE"
End Sub

Private Sub Player1_onLoadNext()
    LoadMP3_2
End Sub

Private Sub Player1_onPause()
    lblPlayer1title.Caption = "PLAYER A: READY"
End Sub

Private Sub Player1_onPlay()
    Timer1.Enabled = True
    Player1Tmr = 0
    CrossFader_ValueChanged
    lblPlayer1title.Caption = "PLAYER A: PLAYING"
    IsPlaying = True
    
    If Player1.Mode = pNORMAL Then
        lblStatus.Caption = "Player 1 currently playing " & Trim(Player1.pTitle)
    End If
    
    If Player1.Mode = pBACKGROUND Then
        lblStatus.Caption = "Player 1 is currently background mode, Press Next button to play next song"
    End If
    
End Sub

Private Sub Player1_onStartMix()
    Player2.UnloadMP3
End Sub

Private Sub Player1_onStop()
    Timer1.Enabled = False
    analyzer(0).Cls
    IsPlaying = False
    lblStatus.Caption = "Player 1 stopped"
    lblPlayer1title.Caption = "PLAYER A: STOPPED"
End Sub

Private Sub Player1_onUnloadMP3()
    lblTitle(0).Caption = ""
    lblArtist(0).Caption = ""
    lblAlbum(0).Caption = ""
    lblGenre(0).Caption = ""
    lblLen(0).Caption = "00:00:00"
    lblTrig(0).Caption = "00:00:00"
    lblCurTime(0).Caption = "00:00:00"
    lblRem(0).Caption = "00:00:00"
    analyzer(0).Cls
    pos.max = 100
    pos.Min = 0
    pos.value = 0
    lblPlayer1title.Caption = "PLAYER A: EMPTY"
End Sub

Private Sub Player2_onEndMix()
    Player1.setPosition Player1.pSongStart
    Player1.Play
End Sub

Private Sub Player2_onLoadMp3()
    lblTitle(1).Caption = Player2.pTitle
    lblArtist(1).Caption = Player2.pArtist
    lblAlbum(1).Caption = Player2.pAlbum
    lblGenre(1).Caption = Player2.pGenre
    posb.max = Player2.getStreamlenght()
    posb.Min = 0
    lblPlayer2title.Caption = "PLAYER B: IDLE"

End Sub

Private Sub Player2_onLoadNext()
    LoadMP3_1
End Sub

Private Sub Player2_onPause()
    lblPlayer2title.Caption = "PLAYER B: READY"
End Sub

Private Sub Player2_onPlay()
    Timer2.Enabled = True
    Player2Tmr = 0
    CrossFader_ValueChanged
    lblPlayer2title.Caption = "PLAYER B: PLAYING"
    IsPlaying = True
    
    If Player2.Mode = pNORMAL Then
        lblStatus.Caption = "Player 2 currently playing " & Trim(Player2.pTitle)
    End If
    
    If Player2.Mode = pBACKGROUND Then
        lblStatus.Caption = "Player 2 is currently background mode, Press Next button to play next song"
    End If
    
End Sub

Private Sub Player2_onStartMix()
    Player1.UnloadMP3
End Sub

Private Sub Player2_onStop()
    Timer2.Enabled = False
    analyzer(1).Cls
    analyzer(1).Refresh
    lblStatus.Caption = "Player 2 stopped"
    
    IsPlaying = False
    lblPlayer2title.Caption = "PLAYER B: STOPPED"
End Sub

Private Sub Player2_onUnloadMP3()
    lblTitle(1).Caption = ""
    lblArtist(1).Caption = ""
    lblAlbum(1).Caption = ""
    lblGenre(1).Caption = ""
    lblLen(1).Caption = "00:00:00"
    lblTrig(1).Caption = "00:00:00"
    lblCurTime(1).Caption = "00:00:00"
    lblRem(1).Caption = "00:00:00"
    
    posb.max = 100
    posb.Min = 0
    posb.value = 0
    lblPlayer2title.Caption = "PLAYER B: EMPTY"
    analyzer(1).Cls
End Sub

Private Sub pos_MouseDown(Shift As Integer)
    Timer1.Enabled = False
End Sub

Private Sub pos_MouseUp(Shift As Integer)
    Player1.setPosition pos.value
    Timer1.Enabled = True
End Sub

Private Sub posb_MouseDown(Shift As Integer)
    Timer2.Enabled = False
End Sub

Private Sub posb_MouseUp(Shift As Integer)
    Player2.setPosition posb.value
    Timer2.Enabled = True
End Sub



Private Sub samp_Click(Index As Integer)
Select Case Index
Case Is = 0
    BASS_SamplePlayEx Sample1, 0, -1, 100, Int((201 * Rnd) - 100), BASSFALSE
Case Is = 1
    BASS_SamplePlayEx Sample2, 0, -1, 100, Int((201 * Rnd) - 100), BASSFALSE
Case Is = 2
    BASS_SamplePlayEx Sample3, 0, -1, 100, Int((201 * Rnd) - 100), BASSFALSE
Case Is = 3
    BASS_SamplePlayEx Sample4, 0, -1, 100, Int((201 * Rnd) - 100), BASSFALSE
End Select

End Sub

Private Sub StopDeck_Click()
'StopMP3
Player1.StopPlay
End Sub

Public Sub Timer1_Timer()
    
If Player1.IsPlaying = True Then
    pos.value = Player1.getPosition()
    lblCurTime(0).Caption = Player1.getCurrentTime()
    lblRem(0).Caption = GetTime(DateDiff("s", lblCurTime(0).Caption, lblTrig(0).Caption))
        
    If Player2MixNext = True And Player1.Mode = pNORMAL Then
        Player2.Play
        Player1.StopPlay
        Player2MixNext = False
    End If

        
    If Player1.Mode = pBACKGROUND And Player2MixNext = True Then
        Player1.setPosition Player1.pSongStart
        Player2MixNext = False
    End If




    If Player1.AudioLevel = 0 Then
        Player1.fxAGC = False
        Player1.fxAGC = True
    End If

    'display spectrum
    Player1.getSpectrum analyzer(0), abuff, gph

End If

    Player1.getSpectrum analyzer(0), abuff, gph

End Sub

Private Function Sqroot(ByVal num As Double) As Double
    Sqroot = num ^ 0.5
End Function


Private Sub Timer2_Timer()
    If Player2.IsPlaying = True Then
        posb.value = Player2.getPosition()
        lblCurTime(1).Caption = Player2.getCurrentTime()
        lblRem(1).Caption = GetTime(DateDiff("s", lblCurTime(1).Caption, lblTrig(1).Caption))
        
        If Player1MixNext = True And Player2.Mode = pNORMAL Then
            Player1.Play
            Player2.StopPlay
            Player1MixNext = False
        End If
            
        If Player2.Mode = pBACKGROUND And Player1MixNext = True Then
            Player2.setPosition Player2.pSongStart
            Player1MixNext = False
        End If
            
            
        If Player2.AudioLevel = 0 Then
            Player2.fxAGC = False
            Player2.fxAGC = True
        End If
        
    End If

   
   Player2.getSpectrum analyzer(1), abuff, gph

End Sub

Private Sub tmrClock_Timer()
    lblTime.Caption = time
    lblMemory.Caption = Format(BASS_GetCPU(), "0.0")
    
    If Player1.IsPlaying = True Or Player2.IsPlaying = True Then
        If TurnTableIsActive = False Then
            DAControl.Start
            TurnTableIsActive = True
        End If
    Else
        If TurnTableIsActive = True Then
            DAControl.Stop
            TurnTableIsActive = False
        End If
    End If
    
    
    If Player1.IsPlaying = False And Player2.IsPlaying = False Then
        UpdateTimePlay
    End If
End Sub


'============================================

Private Sub LoadUserPlaylist()
    Dim lstcount As Integer
    Dim i As Integer
    Dim idx As Integer
    Dim plbuff() As Long
    
    Dim bSize As Long
    
    Open App.Path & "\user.upl" For Binary As #2
    Get #2, 1, bSize          ' Read the number of items.
    ReDim plbuff(0 To bSize) As Long
    Get #2, , plbuff()
    Close #2
   
    For i = 1 To UBound(plbuff)
        idx = plbuff(i)
        LoadPlayList idx
     Next i
    
End Sub

Public Sub LoadPlayList(idx As Integer)
       If idx <> 0 Then
       If idx > UBound(MUSIC) Then Exit Sub
       
       If FileExist(MUSIC(idx).file) = False Then Exit Sub
       
       Set lst = ListView1.ListItems.Add()
       'Lst.SmallIcon = 1
       lst.text = MUSIC(idx).Len
       lst.SubItems(1) = MUSIC(idx).Title
       lst.SubItems(2) = MUSIC(idx).Artist
       lst.SubItems(3) = Format(MUSIC(idx).bpm, "000")
       lst.SubItems(4) = idx
       lst.SubItems(5) = MUSIC(idx).Genre
       lst.SubItems(6) = MUSIC(idx).Year
       lst.SubItems(7) = MUSIC(idx).Rating
       lst.SubItems(8) = MUSIC(idx).LastPlay
       End If
End Sub

Private Sub SaveCurrentList()
    Dim lstcount As Integer
    Dim i As Integer
    Dim idx As Integer
    Dim plbuff() As Long
    
    lstcount = ListView1.ListItems.count
    If lstcount = 0 Then
        MsgBox "Playlist is empty", vbOKOnly, "Save error"
        Exit Sub
    End If
    
    ReDim plbuff(lstcount) As Long
    
    For i = 1 To lstcount
        plbuff(i) = ListView1.ListItems(i).SubItems(4)
    Next i
    
    Open App.Path & "\user.upl" For Binary As #1
    Put #1, 1, CLng(UBound(plbuff))
    Put #1, , plbuff()
    Close #1
    
    MsgBox "Playlist saved", vbOKOnly, "Done"
End Sub



Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'If Button = 1 Then
'    ListView1.Drag
'End If
End Sub


Private Sub RemoveFromList()
Dim R As Integer
R = DragSourceTrack
If R = 0 Then R = 1
 
If Playlist.Rows = 2 Or Playlist.Rows = 1 Then
    Playlist.Rows = 1
Else
    Playlist.RemoveItem R
End If
End Sub



Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim pfile As String
    Dim idx As Long
    Dim sqltxt As String

'delete key
If KeyCode = 46 And ListView2.ListItems.count > 0 Then
    idx = ListView2.ListItems.Item(ListView2.SelectedItem.Index).SubItems(3)
    sqltxt = "DELETE * FROM music WHERE index=" & idx
    ExecuteQuery sqltxt
    DisplayListByCategory ListView3.SelectedItem.text
End If

'~ has been pressed
If KeyCode = 32 And ListView2.ListItems.count > 0 Then
    idx = ListView2.ListItems.Item(ListView2.SelectedItem.Index).SubItems(3)
    
    Set rspreview = OpenRS("SELECT * FROM music WHERE index=" & idx)
    If rspreview.RecordCount = 0 Then Exit Sub
    pfile = rspreview.Fields!file
    If BASS_ChannelIsActive(previewchan) = BASS_ACTIVE_PLAYING Then BASS_ChannelStop previewchan
    BASS_StreamFree previewchan
    If AUDIO.EnableCueDevice = 1 Then
        Call BASS_SetDevice(AUDIO.CueDevice)
    End If
    previewchan = BASS_StreamCreateFile(BASSFALSE, pfile, 0, 0, 0)
    Call BASS_ChannelPlay(previewchan, BASSTRUE)
End If

'escape key has been pressed
If KeyCode = 27 And ListView2.ListItems.count > 0 Then
    BASS_ChannelStop previewchan
    BASS_StreamFree previewchan
End If

''37 left arrow has been pressed
'If KeyCode = 37 And ListView2.ListItems.Count > 0 Then
'Call mAddToEnd_Click
'End If

End Sub

Private Sub ListView2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then ListView2.Drag
End Sub


Private Sub ListView3_KeyDown(KeyCode As Integer, Shift As Integer)
'del - delete one by one
If KeyCode = 46 And Shift = 0 And ListView3.ListItems.count > 0 Then
ListView3.ListItems.Remove (ListView3.SelectedItem.Index)
End If
End Sub


Private Sub ListView4_KeyDown(KeyCode As Integer, Shift As Integer)
    'delete key
    If KeyCode = 46 And ListView4.ListItems.count > 0 Then
        ListView4.ListItems.Remove (ListView4.SelectedItem.Index)
        UpdateStation STATION
        SaveStation
    End If

End Sub

Private Sub mClearAll_Click()
    ListView1.ListItems.Clear
End Sub

Private Sub mCopyAllToPlaylist_Click()
    Dim i As Integer
    Dim mindex As Long
    Dim lst As ListItem
     For i = 1 To ListView2.ListItems.count
        mindex = ListView2.ListItems.Item(i).SubItems(3)
        CopyToPlayList mindex
    Next i
    UpdateTimePlay

End Sub

Private Sub mCut_Click()
    Dim idx As Long
    idx = ListView1.ListItems.Item(ListView1.SelectedItem.Index).SubItems(3)
    If idx = -1 Then Exit Sub
    DeleteMusic idx
    
End Sub


Private Sub UpdateStation(rStation() As IRADIO)
'Dim c As Integer
'If ListView4.ListItems.Count = 0 Then Exit Sub
'ReDim rStation(ListView4.ListItems.Count) As IRADIO
'For c = 1 To ListView4.ListItems.Count
'    With rStation(c)
'               .StationName = ListView4.ListItems(c).text
'               .Address = ListView4.ListItems(c).SubItems(1)
'               .Category = ListView4.ListItems(c).SubItems(2)
'    End With
'Next c
End Sub



Private Sub setRating(Index As Integer, cat As String)
    Dim i As Integer
    For i = 0 To UBound(MUSIC())
        If i = Index And MUSIC(i).Rating > 0 Then
            MUSIC(i).Rating = MUSIC(i).Rating - 1
            GoTo SKIP
        End If
SKIP:
    Next i
End Sub

Private Sub decrementRating()
    Dim i As Integer
        'this must be called every end if the session so that those song that does
        'not played anymore must reduce the rating down
    For i = 0 To UBound(MUSIC())
            If MUSIC(i).Rating < 10 Then
             MUSIC(i).Rating = MUSIC(i).Rating + 1
            End If
    Next i
End Sub

Public Function UpdateTimePlay()
    If ListView1.ListItems.count = 0 Then
        lblTotalTime.Caption = "00:00:00"
        Exit Function
    End If
    
    Dim i As Integer
    tSeconds = 0

    curTime = time
    tSeconds = 0
    For i = 1 To ListView1.ListItems.count
        ListView1.ListItems(i).text = Format(DateAdd("s", tSeconds, curTime), "hh:mm:ss")
        tSeconds = tSeconds + Val(ListView1.ListItems(i).SubItems(9))
    Next i
    
    lblTotalTime.Caption = GetTime(tSeconds)
End Function

Public Sub ReloadHistory()
    Dim lstv As ListItem
    Dim idx As Long
    ListView1.ListItems.Clear
    
    Set rsloadhistory = OpenRS("SELECT * FROM history")
    If rsloadhistory.RecordCount <> 0 Then
        Do While Not rsloadhistory.EOF
        idx = rsloadhistory.Fields!MusicIndex
        
        Set rstmpmusic = OpenRS("SELECT * FROM music WHERE index=" & idx)
        If rstmpmusic.RecordCount <> 0 Then
        If FileExist(rstmpmusic.Fields!file) Then
            Set lstv = ListView1.ListItems.Add
            lstv.SubItems(1) = rstmpmusic.Fields!Title
            lstv.SubItems(2) = rstmpmusic.Fields!Artist
            lstv.SubItems(3) = rstmpmusic.Fields!bpm
            lstv.SubItems(4) = rstmpmusic.Fields!Index
            lstv.SubItems(5) = rstmpmusic.Fields!Genre
            lstv.SubItems(6) = rstmpmusic.Fields!SongYear
            lstv.SubItems(7) = rstmpmusic.Fields!Rating
            lstv.SmallIcon = SetIcon(rstmpmusic.Fields!Category)
            
            End If
        End If
        rsloadhistory.MoveNext
        Loop
        
        ExecuteQuery "DELETE * FROM history"
    End If
    
        UpdateTimePlay
    
End Sub


Public Sub LoadCompose(cname As String)
    Dim lstv As ListItem
    Dim idx As Long
    
    Set rsloadhistory = OpenRS("SELECT * FROM composition WHERE composename='" & cname & "'")
    If rsloadhistory.RecordCount <> 0 Then
        Do While Not rsloadhistory.EOF
        idx = rsloadhistory.Fields!MusicIndex
        
        Set rstmpmusic = OpenRS("SELECT * FROM music WHERE index=" & idx)
        If rstmpmusic.RecordCount <> 0 Then
        If FileExist(rstmpmusic.Fields!file) Then
            Set lstv = ListView1.ListItems.Add
            lstv.SubItems(1) = rstmpmusic.Fields!Title
            lstv.SubItems(2) = rstmpmusic.Fields!Artist
            lstv.SubItems(3) = rstmpmusic.Fields!bpm
            lstv.SubItems(4) = rstmpmusic.Fields!Index
            lstv.SubItems(5) = rstmpmusic.Fields!Genre
            lstv.SubItems(6) = rstmpmusic.Fields!SongYear
            lstv.SubItems(7) = rstmpmusic.Fields!Rating
            lstv.SmallIcon = SetIcon(rstmpmusic.Fields!Category)
            
            End If
        End If
        rsloadhistory.MoveNext
        Loop
    End If
        
    UpdateTimePlay
    
End Sub

Public Function MixNow()
        'if both are playing then leave it
        If Player1.IsPlaying = True And Player2.IsPlaying = True Then Exit Function
        
        If Player1.IsPlaying = True And Player2.IsPlaying = False Then
            If Player2.IsLoaded = False Then LoadMP3_2
            Player2.setPosition Player2.pSongStart
            Player2.Play
            Player1.StopPlay
            Exit Function
        End If
        
        If Player2.IsPlaying = True And Player1.IsPlaying = False Then
            If Player1.IsLoaded = False Then LoadMP3_1
            Player1.setPosition Player1.pSongStart
            Player1.Play
            Player2.StopPlay
            Exit Function
        End If

        
        If Player1.IsPlaying = False And Player2.IsPlaying = False Then
            LoadMP3_1
            Player1.Play
            Exit Function
        End If

End Function


Public Sub DisplayFolderList()
    Dim sqltxt As String
    Dim lst1 As ListItem
    ListView3.ListItems.Clear
    
    
    Set lst1 = ListView3.ListItems.Add
    lst1.text = "~Advertisement"
    lst1.SmallIcon = 1
    
    Set lst1 = ListView3.ListItems.Add
    lst1.text = "~Jingles"
    lst1.SmallIcon = 1
    
    Set lst1 = ListView3.ListItems.Add
    lst1.text = "~Program"
    lst1.SmallIcon = 1
    
    Set lst1 = ListView3.ListItems.Add
    lst1.text = "~StationID"
    lst1.SmallIcon = 1
    
    Set lst1 = ListView3.ListItems.Add
    lst1.text = "~VoiceOverBackground"
    lst1.SmallIcon = 1
    
'    Set lst1 = ListView3.ListItems.Add
'    lst1.text = "~All Music"
'    lst1.SmallIcon = 1
    
    Set rsMusic = OpenRS("SELECT * FROM folder ORDER BY folder")
    Do While Not rsMusic.EOF
            Set lst1 = ListView3.ListItems.Add
            lst1.text = rsMusic.Fields!folder
            lst1.SmallIcon = 2
            rsMusic.MoveNext
    Loop
End Sub


Private Sub DisplayListByCategory(cat As String)
    Dim lstv As ListItem
    ListView2.ListItems.Clear
    
    Me.MousePointer = vbHourglass
    
    
    If cat = "~All Music" Then
        Set rsMusic = OpenRS("SELECT * FROM music ORDER BY Title")
    Else
        Set rsMusic = OpenRS("SELECT * FROM music WHERE Category='" & cat & "'")
    End If
    
    
    If rsMusic.RecordCount <> 0 Then
        Do While Not rsMusic.EOF
        
        If FileExist(rsMusic.Fields!file) Then
        Set lstv = ListView2.ListItems.Add
        lstv.text = rsMusic.Fields!Title
        lstv.SubItems(1) = rsMusic.Fields!Artist
        lstv.SubItems(2) = GetTime(rsMusic.Fields!Len)
        lstv.SubItems(3) = rsMusic.Fields!Index
        lstv.SubItems(4) = rsMusic.Fields!Genre
        lstv.SubItems(5) = rsMusic.Fields!SongYear
        lstv.SubItems(6) = rsMusic.Fields!Rating
        lstv.SubItems(7) = rsMusic.Fields!LastPlay
        lstv.SubItems(8) = rsMusic.Fields!Len
        lstv.SmallIcon = SetIcon(rsMusic.Fields!Category)
        
        End If
        rsMusic.MoveNext
        Loop
    End If
    
    Me.MousePointer = vbNormal
End Sub

Private Sub DisplayPlaylist()
    Dim lstv As ListItem
    Dim idx As Long
    ListView1.ListItems.Clear
    
    Me.MousePointer = vbHourglass
    
    Set rsPlaylist = OpenRS("SELECT * FROM playlist")
    If rsPlaylist.RecordCount <> 0 Then
        Do While Not rsPlaylist.EOF
        idx = rsPlaylist.Fields!MusicIndex
        Set rsMusic = Nothing
        Set rsMusic = OpenRS("SELECT * FROM music WHERE index=" & idx)
        If rsMusic.RecordCount <> 0 Then
            If FileExist(rsMusic.Fields!file) Then
            Set lstv = ListView1.ListItems.Add
            lstv.text = ""
            lstv.SubItems(1) = rsMusic.Fields!Title
            lstv.SubItems(2) = rsMusic.Fields!Artist
            lstv.SubItems(3) = rsMusic.Fields!bpm
            lstv.SubItems(4) = rsMusic.Fields!Index
            lstv.SubItems(5) = rsMusic.Fields!Genre
            lstv.SubItems(6) = rsMusic.Fields!SongYear
            lstv.SubItems(7) = rsMusic.Fields!Rating
            lstv.SubItems(8) = rsMusic.Fields!LastPlay
            End If
        End If
        rsPlaylist.MoveNext
        Loop
    End If
    
    Me.MousePointer = vbNormal
End Sub

Private Sub ClsPlaylist()
    ExecuteQuery "DELETE FROM playlist"
    DisplayPlaylist
End Sub



Private Sub SavePlayList()
    Dim i As Integer
    Dim sqltxt As String
    
    Me.MousePointer = vbHourglass
    
    For i = 1 To UBound(MUSIC)
        sqltxt = "INSERT INTO music(file,title,artist,SongStart,SongEnd,bpm,Len,Category,LastPlay,MixType,Rating,Genre,Album,SongYear) "
        sqltxt = sqltxt & " VALUES('" & QuoteReplace(MUSIC(i).file) & " ',"
        sqltxt = sqltxt & "'" & QuoteReplace(MUSIC(i).Title) & " ',"
        sqltxt = sqltxt & "'" & QuoteReplace(MUSIC(i).Artist) & " ',"
        sqltxt = sqltxt & MUSIC(i).SongStart & ","
        sqltxt = sqltxt & MUSIC(i).SongEnd & ","
        sqltxt = sqltxt & MUSIC(i).bpm & ","
        sqltxt = sqltxt & MUSIC(i).Len & ","
        sqltxt = sqltxt & "'" & MUSIC(i).Category & " ',"
        sqltxt = sqltxt & "'" & MUSIC(i).LastPlay & " ',"
        sqltxt = sqltxt & "'" & MUSIC(i).MixType & " ',"
        sqltxt = sqltxt & MUSIC(i).Rating & ","
        sqltxt = sqltxt & "'" & MUSIC(i).Genre & " ',"
        sqltxt = sqltxt & "'" & QuoteReplace(MUSIC(i).Album) & " ',"
        sqltxt = sqltxt & "'" & QuoteReplace(MUSIC(i).Year) & " ')"
        ExecuteQuery sqltxt
   Next i
   
   Me.MousePointer = vbNormal
End Sub


Private Sub SetEqualizer()
    If eqEnable.Status = True Then
        Dim i As Integer
        Player1.EqualizerEnable False
        Player2.EqualizerEnable False
        
        If Player1.IsLoaded = True Then
            Player1.clsEqualizer.eBand1 = Equalizer1(0).Frequency
            Player1.clsEqualizer.eBand2 = Equalizer1(1).Frequency
            Player1.clsEqualizer.eBand3 = Equalizer1(2).Frequency
            Player1.clsEqualizer.eBand4 = Equalizer1(3).Frequency
            Player1.clsEqualizer.eBand5 = Equalizer1(4).Frequency
            Player1.clsEqualizer.eBand6 = Equalizer1(5).Frequency
            Player1.clsEqualizer.eBand7 = Equalizer1(6).Frequency
            Player1.clsEqualizer.eBandWidth = 1.5
            Player1.EqualizerEnable True
            Player1.clsEqualizer.egain = 2
            For i = 0 To 6
               Player1.UpdateEqualizer i, (Equalizer1(i).value - 7) * -1
            Next i
        End If
        
        If Player2.IsLoaded = True Then
            Player2.clsEqualizer.eBand1 = Equalizer1(0).Frequency
            Player2.clsEqualizer.eBand2 = Equalizer1(1).Frequency
            Player2.clsEqualizer.eBand3 = Equalizer1(2).Frequency
            Player2.clsEqualizer.eBand4 = Equalizer1(3).Frequency
            Player2.clsEqualizer.eBand5 = Equalizer1(4).Frequency
            Player2.clsEqualizer.eBand6 = Equalizer1(5).Frequency
            Player2.clsEqualizer.eBand7 = Equalizer1(6).Frequency
            Player2.clsEqualizer.eBandWidth = 1.5
            Player2.EqualizerEnable True
            Player2.clsEqualizer.egain = 2
            For i = 0 To 6
               Player2.UpdateEqualizer i, (Equalizer1(i).value - 7) * -1
            Next i
        End If
        
    Else
    
        Player1.EqualizerEnable False
        Player2.EqualizerEnable False
        
    End If
End Sub

Private Sub PlayLoop(Index As Integer, stat As Boolean)
Select Case Index
Case Is = 0
    If stat = True Then
        pfile = App.Path & "\loops\loop1.wav"
        If FileExist(pfile) = False Then Exit Sub
        'kung naga tokar hay e stop
        If BASS_ChannelIsActive(LoopChan) = BASS_ACTIVE_PLAYING Then BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
        LoopChan = BASS_StreamCreateFile(BASSFALSE, pfile, 0, 0, BASS_SAMPLE_LOOP)
      'set pitch
        Call BASS_ChannelGetAttributes(LoopChan, freq, vbNull, vbNull)
        With loopspeed
        .max = (freq * maxp) / 4
        .Min = (freq * minp) / 4
        .value = freq / 4
        End With
        Call BASS_ChannelPlay(LoopChan, 0, BASS_SAMPLE_LOOP)
    Else
        BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
    End If
Case Is = 1
    If stat = True Then
        pfile = App.Path & "\loops\loop2.wav"
        If FileExist(pfile) = False Then Exit Sub
        'kung naga tokar hay e stop
        If BASS_ChannelIsActive(LoopChan) = BASS_ACTIVE_PLAYING Then BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
        LoopChan = BASS_StreamCreateFile(BASSFALSE, pfile, 0, 0, BASS_SAMPLE_LOOP)
      'set pitch
        Call BASS_ChannelGetAttributes(LoopChan, freq, vbNull, vbNull)
        With loopspeed
        .max = (freq * maxp) / 4
        .Min = (freq * minp) / 4
        .value = freq / 4
        End With
        Call BASS_ChannelPlay(LoopChan, 0, BASS_SAMPLE_LOOP)
    Else
        BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
    End If
Case Is = 2
    If stat = True Then
        pfile = App.Path & "\loops\loop3.wav"
        If FileExist(pfile) = False Then Exit Sub
        'kung naga tokar hay e stop
        If BASS_ChannelIsActive(LoopChan) = BASS_ACTIVE_PLAYING Then BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
        LoopChan = BASS_StreamCreateFile(BASSFALSE, pfile, 0, 0, BASS_SAMPLE_LOOP)
      'set pitch
        Call BASS_ChannelGetAttributes(LoopChan, freq, vbNull, vbNull)
        With loopspeed
        .max = (freq * maxp) / 4
        .Min = (freq * minp) / 4
        .value = freq / 4
        End With
        Call BASS_ChannelPlay(LoopChan, 0, BASS_SAMPLE_LOOP)
    Else
        BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
    End If
Case Is = 3
    If stat = True Then
        pfile = App.Path & "\loops\loop4.wav"
        If FileExist(pfile) = False Then Exit Sub
        'kung naga tokar hay e stop
        If BASS_ChannelIsActive(LoopChan) = BASS_ACTIVE_PLAYING Then BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
        LoopChan = BASS_StreamCreateFile(BASSFALSE, pfile, 0, 0, BASS_SAMPLE_LOOP)
      'set pitch
        Call BASS_ChannelGetAttributes(LoopChan, freq, vbNull, vbNull)
        With loopspeed
        .max = (freq * maxp) / 4
        .Min = (freq * minp) / 4
        .value = freq / 4
        End With
        Call BASS_ChannelPlay(LoopChan, 0, BASS_SAMPLE_LOOP)
    Else
        BASS_ChannelStop LoopChan
        BASS_StreamFree LoopChan
    End If
End Select

End Sub

Private Sub ClearHistory()
    ExecuteQuery "DELETE * FROM history"
End Sub


Private Sub TmrVol_Timer()
    If TalkActive Then
        If ledSmoth.Status = True Then
            volMaster.value = volMaster.value - 5
        Else
            volMaster.value = TalkVol
            TmrVol.Enabled = False
        End If
        If volMaster.value <= TalkVol Then
            volMaster.value = TalkVol
            TmrVol.Enabled = False
        End If
    Else
        If ledSmoth.Status = False Then
            volMaster.value = OrigVol
            TmrVol.Enabled = False
        Else
            volMaster.value = volMaster.value + 5
        End If
        
        If volMaster.value > OrigVol Then
            volMaster.value = OrigVol
            TmrVol.Enabled = False
        End If
    End If

End Sub

Private Sub volMaster_MouseUp(Shift As Integer)
    'if the Talk mode and they adjust the volume set this as new talkvol value
    If TalkActive Then
        TalkVol = volMaster.value
    End If

End Sub

Private Sub volMaster_ValueChanged()
    modBass.BASS_SetVolume volMaster.value
    
End Sub
