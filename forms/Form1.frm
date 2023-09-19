VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Player1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8685
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E49667&
      BorderStyle     =   0  'None
      Height          =   5580
      Left            =   0
      ScaleHeight     =   5580
      ScaleWidth      =   11970
      TabIndex        =   35
      Top             =   3105
      Width           =   11970
      Begin VB.PictureBox Picture14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4515
         Left            =   0
         ScaleHeight     =   4515
         ScaleWidth      =   6105
         TabIndex        =   36
         Top             =   1065
         Width           =   6105
         Begin MSComctlLib.ListView ListView2 
            Height          =   4230
            Left            =   15
            TabIndex        =   37
            Top             =   270
            Width           =   6075
            _ExtentX        =   10716
            _ExtentY        =   7461
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   16777215
            BackColor       =   10638871
            Appearance      =   0
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
               Text            =   "Play Time"
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
         End
         Begin VB.Label lblTotalTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00:00:00 - Total time"
            Height          =   195
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   1455
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5580
         Left            =   6105
         TabIndex        =   39
         Top             =   0
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   9843
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         TabHeight       =   2
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "Form1.frx":630A2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblMediaLibraryTitle"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Picture3"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Tab 1"
         TabPicture(1)   =   "Form1.frx":630BE
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "Form1.frx":630DA
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.PictureBox picRadioBox 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4275
            Left            =   -74940
            ScaleHeight     =   4275
            ScaleWidth      =   5745
            TabIndex        =   52
            Top             =   480
            Width           =   5745
            Begin VB.CheckBox chkRecord 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Record "
               Enabled         =   0   'False
               Height          =   300
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   57
               Top             =   375
               Width           =   825
            End
            Begin VB.CommandButton cmdConnect 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Connect"
               Height          =   315
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   56
               ToolTipText     =   "Connect to Radion and Play"
               Top             =   30
               Width           =   825
            End
            Begin VB.PictureBox picRadioLcd 
               BackColor       =   &H00F8C4A7&
               Height          =   1245
               Left            =   870
               ScaleHeight     =   1185
               ScaleWidth      =   4770
               TabIndex        =   54
               Top             =   30
               Width           =   4830
               Begin VB.Label lblRadioStatus 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Idle..."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   900
                  Left            =   105
                  TabIndex        =   55
                  Top             =   75
                  Width           =   4575
               End
            End
            Begin VB.CommandButton cmdAddStation 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Add Sta."
               Height          =   315
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   53
               ToolTipText     =   "Add station address"
               Top             =   735
               Width           =   825
            End
            Begin MSComctlLib.ListView ListView4 
               Height          =   2130
               Left            =   45
               TabIndex        =   58
               Top             =   1320
               Width           =   5685
               _ExtentX        =   10028
               _ExtentY        =   3757
               View            =   3
               Arrange         =   2
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               _Version        =   393217
               SmallIcons      =   "ImageList1"
               ForeColor       =   16777215
               BackColor       =   14980711
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Station Name"
                  Object.Width           =   5292
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Address"
                  Object.Width           =   4410
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Category"
                  Object.Width           =   1764
               EndProperty
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5010
            Left            =   60
            ScaleHeight     =   5010
            ScaleWidth      =   5745
            TabIndex        =   40
            Top             =   510
            Width           =   5745
            Begin VB.OptionButton cmdPL 
               BackColor       =   &H00E0E0E0&
               Caption         =   "All Album"
               ForeColor       =   &H00404000&
               Height          =   270
               Index           =   0
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   15
               Width           =   1140
            End
            Begin VB.OptionButton cmdPL 
               Caption         =   "1980s"
               Height          =   270
               Index           =   1
               Left            =   1155
               Style           =   1  'Graphical
               TabIndex        =   49
               Top             =   15
               Width           =   1140
            End
            Begin VB.OptionButton cmdPL 
               Caption         =   "1990s"
               Height          =   270
               Index           =   2
               Left            =   2310
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   15
               Width           =   1140
            End
            Begin VB.OptionButton cmdPL 
               Caption         =   "2000s"
               Height          =   270
               Index           =   3
               Left            =   3465
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   15
               Width           =   1140
            End
            Begin VB.OptionButton cmdPL 
               Caption         =   "Dance"
               Height          =   270
               Index           =   4
               Left            =   4605
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   15
               Width           =   1140
            End
            Begin VB.OptionButton cmdPL 
               Caption         =   "Favorites"
               Height          =   270
               Index           =   5
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   285
               Width           =   1140
            End
            Begin VB.OptionButton cmdPL 
               Caption         =   "Rock"
               Height          =   270
               Index           =   6
               Left            =   1155
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   285
               Width           =   1140
            End
            Begin VB.OptionButton cmdPL 
               Caption         =   "Station ID"
               Height          =   270
               Index           =   7
               Left            =   2310
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   285
               Width           =   1140
            End
            Begin VB.OptionButton cmdPL 
               Caption         =   "Advertise"
               Height          =   270
               Index           =   8
               Left            =   3465
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   285
               Width           =   1140
            End
            Begin VB.OptionButton cmdPL 
               Caption         =   "Jingle"
               Height          =   270
               Index           =   9
               Left            =   4620
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   285
               Width           =   1125
            End
            Begin MSComctlLib.ImageList ImageList1 
               Left            =   750
               Top             =   2670
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   16
               ImageHeight     =   16
               MaskColor       =   12632256
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   6
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Form1.frx":630F6
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Form1.frx":6354A
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Form1.frx":6399E
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Form1.frx":64096
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Form1.frx":6476A
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Form1.frx":64BBE
                     Key             =   ""
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   4410
               Left            =   15
               TabIndex        =   51
               Top             =   570
               Width           =   5700
               _ExtentX        =   10054
               _ExtentY        =   7779
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               HoverSelection  =   -1  'True
               _Version        =   393217
               SmallIcons      =   "ImageList1"
               ForeColor       =   0
               BackColor       =   14326369
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Title"
                  Object.Width           =   5292
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Artist"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "BPM"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "ID"
                  Object.Width           =   9
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Genre"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Year"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Rating"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "Lastplay"
                  Object.Width           =   3528
               EndProperty
            End
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   4080
            Left            =   -74850
            TabIndex        =   59
            Top             =   585
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   7197
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   0
            BackColor       =   14737632
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Title"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Artist"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "BPM"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "ID"
               Object.Width           =   353
            EndProperty
         End
         Begin VB.Label lblPlaylistTitle 
            Caption         =   "Playlist History list"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00F8C4A7&
            Height          =   390
            Left            =   -74955
            TabIndex        =   62
            Top             =   45
            Width           =   2970
         End
         Begin VB.Label lblInternetRadioTitle 
            Caption         =   "Internet Radio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00F8C4A7&
            Height          =   315
            Left            =   -74880
            TabIndex        =   61
            Top             =   30
            Width           =   2655
         End
         Begin VB.Label lblMediaLibraryTitle 
            Caption         =   "Media Library"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00F8C4A7&
            Height          =   315
            Left            =   75
            TabIndex        =   60
            Top             =   15
            Width           =   2655
         End
      End
   End
   Begin VB.CommandButton fx1 
      Caption         =   "Reverb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2190
      Width           =   570
   End
   Begin VB.CommandButton fx2 
      Caption         =   "Flanger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2505
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2190
      Width           =   570
   End
   Begin VB.CommandButton fx3 
      Caption         =   "Echo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2190
      Width           =   570
   End
   Begin VB.CommandButton fx4 
      Caption         =   "H-pitch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3675
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2190
      Width           =   570
   End
   Begin VB.CommandButton cmdResetBPM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4350
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1770
      Width           =   240
   End
   Begin VB.PictureBox abuff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E8B479&
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
      Height          =   1395
      Left            =   45
      ScaleHeight     =   1365
      ScaleWidth      =   2685
      TabIndex        =   23
      Top             =   3900
      Width           =   2715
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8B479&
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
      Height          =   1605
      Left            =   300
      ScaleHeight     =   1605
      ScaleWidth      =   3810
      TabIndex        =   7
      Top             =   330
      Width           =   3810
      Begin VB.PictureBox analyzer 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8B479&
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
         Left            =   45
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   174
         TabIndex        =   8
         Top             =   15
         Width           =   2610
      End
      Begin VB.Label lblGenre 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
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
         Height          =   225
         Left            =   2655
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   1530
         Width           =   360
      End
      Begin VB.Label lblAlbum 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
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
         Left            =   540
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   1500
         Width           =   435
      End
      Begin VB.Label Player_A 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
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
         TabIndex        =   22
         Top             =   1245
         Width           =   3120
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   1065
         Width           =   345
      End
      Begin VB.Label Player_A 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2985
         TabIndex        =   19
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Player_A 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2985
         TabIndex        =   18
         Top             =   0
         Width           =   675
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   690
         Width           =   195
      End
      Begin VB.Label Player_A 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2985
         TabIndex        =   13
         Top             =   630
         Width           =   675
      End
      Begin VB.Label Player_A 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   2985
         TabIndex        =   12
         Top             =   210
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
         Index           =   6
         Left            =   2715
         TabIndex        =   11
         Top             =   885
         Width           =   195
      End
      Begin VB.Label lblBPM 
         Alignment       =   2  'Center
         BackColor       =   &H00E8B479&
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   2985
         TabIndex        =   10
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Player_A 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8B479&
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
         TabIndex        =   9
         Top             =   1035
         Width           =   3120
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4740
      Top             =   195
   End
   Begin VB.PictureBox bpmbuff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
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
      Height          =   135
      Left            =   105
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   6
      Top             =   2595
      Width           =   135
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
      Height          =   1485
      Left            =   2820
      ScaleHeight     =   1455
      ScaleWidth      =   180
      TabIndex        =   4
      Top             =   4050
      Width           =   210
   End
   Begin Whoosh.cpvSlider pitchSlider 
      Height          =   1380
      Left            =   4380
      Top             =   300
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   423
      BackColor       =   4210688
      SliderIcon      =   "Form1.frx":64EDA
      RailPicture     =   "Form1.frx":6516C
      RailStyle       =   99
      Value           =   5
   End
   Begin Whoosh.cpvSlider pos 
      Height          =   120
      Left            =   300
      Top             =   1950
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   212
      BackColor       =   16303271
      SliderIcon      =   "Form1.frx":65188
      Orientation     =   0
      RailPicture     =   "Form1.frx":6565A
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
      Picture         =   "Form1.frx":65676
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
      Left            =   765
      TabIndex        =   2
      ToolTipText     =   "Play loaded track"
      Top             =   2160
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      Caption         =   ""
      Picture         =   "Form1.frx":65EB0
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
      Picture         =   "Form1.frx":666EA
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
   Begin Whoosh.GurhanButton fx5 
      Height          =   150
      Left            =   4320
      TabIndex        =   5
      ToolTipText     =   "Dynamic gain "
      Top             =   2175
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   265
      Caption         =   ""
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   0   'False
      Raised          =   -1  'True
      ForeColor       =   16711680
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   5145
      Top             =   210
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
            Picture         =   "Form1.frx":66F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":67378
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":677CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":67C20
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":68074
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":684C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6891C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":68D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":691C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":69618
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":69A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":69EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6A314
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6A768
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6ABBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6B010
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   540
      Left            =   0
      TabIndex        =   34
      Top             =   2565
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   953
      ButtonWidth     =   1032
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Play"
            Key             =   "play"
            Object.ToolTipText     =   "Play loaded song"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pause"
            Key             =   "pause"
            Object.ToolTipText     =   "Pause current playing song"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "stop"
            Object.ToolTipText     =   "Stop current playing song"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next"
            Key             =   "next"
            Object.ToolTipText     =   "Fade to next song"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Auto"
            Key             =   "autodj"
            Object.ToolTipText     =   "Auto DJ"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Equ."
            Key             =   "equalizer"
            Object.ToolTipText     =   "Show Equalizer"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mic"
            Key             =   "mic"
            Object.ToolTipText     =   "Enable/Disable mic"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Audio"
            Key             =   "audio"
            Object.ToolTipText     =   "Audio control"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Brows."
            Key             =   "browse"
            Object.ToolTipText     =   "Browse Mp3 files"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mixer"
            Key             =   "mixer"
            Object.ToolTipText     =   "Hide/Show Crossfader"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "save"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Load"
            Key             =   "load"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Wiz."
            Key             =   "wizard"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lib."
            Key             =   "library"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Radio"
            Key             =   "radio"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hist."
            Key             =   "history"
            ImageIndex      =   16
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   195
      Left            =   4290
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AGC"
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
      Height          =   150
      Left            =   4320
      TabIndex        =   24
      Top             =   1965
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   1440
      Left            =   4350
      Shape           =   4  'Rounded Rectangle
      Top             =   270
      Width           =   240
   End
   Begin VB.Label Label2 
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
      Top             =   105
      Width           =   1860
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Player1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const maxp = 1.6
Const minp = 0.4
Dim freq As Long
Dim vol As Integer
Dim LoadPoint As Long

Dim i As Integer
Dim j As Integer
Dim file As String
Dim SampleTemp(0 To 300) As Single

Public MeReadyToStop As Boolean

Public Function StopMP3()
    'Call BASS_ChannelSetPosition(Mp3(1).chan, Mp3(1).file.SongStart)
    Call BASS_ChannelSlideAttributes(BASS_FX_TempoGetResampledHandle(Mp3(1).chan), -1, -2, -101, 1000)
    
    'Call BASS_FX_TempoStopAndFlush(Mp3(1).chan)
    Mp3(1).Estado = STOPED
    Label2.Caption = "DECK A : STOPED"
    analyzer.Refresh
    Timer1.Enabled = False

End Function

Public Function PauseMP3()
    Call BASS_ChannelPause(BASS_FX_TempoGetResampledHandle(Mp3(1).chan))
    Timer1.Enabled = False
    Mp3(1).Estado = PAUSED
    Label2.Caption = "DECK A : PAUSED"
    
End Function
Public Function UnloadMP3()
    'kung naga tokar hay e estap anay
    If Mp3(1).Estado = PLAYING Then StopMP3
    'kag e dis karga ang mp3
    BASS_StreamFree Mp3(1).chan
    Player_A(0).Caption = ""
    Player_A(1).Caption = ""
    Player_A(2).Caption = "00:00:00"
    Player_A(3).Caption = "00:00:00"
    Player_A(4).Caption = "00:00:00"
    Player_A(5).Caption = "00:00:00"
    lblBPM.Caption = ""
    lblAlbum.Caption = ""
    lblGenre.Caption = ""

    
    Mp3(1).Estado = IDLE
    Label2.Caption = "DECK A : IDLE"
    pos.value = 0
    LoadPoint = 0
    analyzer.Cls
End Function
Public Function LoadMP3()
Dim file As String
Dim length As Long
Dim idx As Integer
    'nothing to load then exit
    If Master.ListView2.ListItems.Count < 1 Then Exit Function
    If Timer1.Enabled = True Then
        Timer1.Enabled = False
    End If
    If IsDrag = False Then PlSourceRow = 1
    
    'load the file from the list
    idx = Master.ListView2.ListItems.Item(PlSourceRow).SubItems(4)
    Master.ListView2.ListItems.Remove (PlSourceRow)
    
    'save time and date last play
    MUSIC(idx).LastPlay = Now
    
    'save loaded song to history list
    Dim Lst As ListItem
    Set Lst = Master.ListView3.ListItems.Add()
       Lst.SmallIcon = 1
       Lst.text = MUSIC(idx).Title
       Lst.SubItems(1) = MUSIC(idx).Artist
       Lst.SubItems(2) = MUSIC(idx).bpm
       Lst.SubItems(3) = idx
    If Master.ListView2.ListItems.Count < 1 Then Master.ReloadHistory
    
    With Mp3(1)
        .file = MUSIC(idx)
        Player_A(0).Caption = .file.Title
        Player_A(1).Caption = .file.Artist
        lblAlbum.Caption = .file.Album
        lblGenre.Caption = .file.Genre
        .orgBPM = .file.bpm
    End With
    
    'buhian ang dati nga stream
    Call BASS_FX_DSP_Remove(Mp3(1).chan, BASS_FX_DSPFXVOLUME)
    Call BASS_FX_DSP_Remove(Mp3(1).chan, BASS_FX_DSPFX_PEAKEQ)
    Call BASS_FX_DSP_Remove(Mp3(1).chan, BASS_FX_DSPFX_FLANGER2)
    Call BASS_FX_BPM_Free(Mp3(1).chan)         'free the callback bpm
    Call BASS_FX_BPM_Free(Mp3(1).bpmhandle)    'free the decoding bpm
    Call BASS_FX_TempoFree(Mp3(1).chan)
    BASS_StreamFree Mp3(1).chan
    

    '----------------MP3-----------------------------
'    Mp3(1).chan = BASS_StreamCreateFile(BASSFALSE, Mp3(1).file.file, 0, 0, BASS_SAMPLE_LOOP Or BASS_STREAM_DECODE)
    Mp3(1).chan = BASS_StreamCreateFile(BASSFALSE, Mp3(1).file.file, 0, 0, BASS_STREAM_DECODE)
    
    Call BASS_ChannelGetAttributes(Mp3(1).chan, freq, vbNull, vbNull)
    Call BASS_FX_TempoCreate(Mp3(1).chan, 0)
      
      '--Retrieve Stream info---
    length = BASS_StreamGetLength(Mp3(1).chan)
    BASS_ChannelSetPosition Mp3(1).chan, length - 1
       
       'set pitch
    With pitchSlider
        .max = 1
        .Min = 0
        .max = (maxp - 1) * 100
        .Min = (minp - 1) * 100
        .value = 0
    End With
    
    'length of music
    Player_A(3).Caption = Atime.GetTime(Mp3(1).file.Len)
    'trigger point time
    Player_A(5).Caption = Atime.GetTime(BASS_ChannelBytes2Seconds(Mp3(1).chan, Mp3(1).file.SongEnd))
    'e preparar ang parametric ekwalayser
    Call EqualizerFrm.EqEnable_Click
    'set the volume
    pos.max = BASS_StreamGetLength(Mp3(1).chan)
    pos.Min = 0
    pos.value = 0
    'get the load point
    LoadPoint = pos.max - (pos.max / 2)
    Mp3(1).Estado = READY
    Label2.Caption = "DECK A : READY"
    'StopMP3 'set the player ready
    Call Master.UpdateTimePlay
    'set the mixing point position
    SNYC1 = BASS_ChannelSetSync(Mp3(1).chan, BASS_SYNC_POS, Mp3(1).file.SongEnd, AddressOf modPublic.EndSync, 1)     ' set end sync
    SNYC1a = BASS_ChannelSetSync(Mp3(1).chan, BASS_SYNC_POS, Mp3(1).file.SongStart + 300000, AddressOf modPublic.UnloadSync, 1)    ' set end sync

    lblBPM.Caption = Mp3(1).orgBPM
    MixOut = False
    Player1.MeReadyToStop = False
    
    
    'lblBPM.Caption = DecodeBPM1(True, 0, 30, Mp3(1).file.file)
    'Call BASS_FX_BPM_CallbackSet(Mp3(1).chan, AddressOf GetBPM_Callback1, 5, 0, BASS_FX_BPM_MULT2)
    Call BASS_ChannelSetPosition(Mp3(1).chan, Mp3(1).file.SongStart)
    Call BASS_StreamPlay(BASS_FX_TempoGetResampledHandle(Mp3(1).chan), 0, BASS_STREAM_AUTOFREE)
    Call BASS_ChannelPause(BASS_FX_TempoGetResampledHandle(Mp3(1).chan))
    
    efx5 False
    efx5 True
    fx5.BackColor = RGB(0, 255, 0)
End Function

Public Function PlayMP3()
'kung wala load endi mag play
If Mp3(1).Estado = IDLE Then Exit Function
Timer1.interval = 5
Timer1.Enabled = True

    Select Case Mp3(1).Estado
    Case READY
        Call BASS_ChannelResume(BASS_FX_TempoGetResampledHandle(Mp3(1).chan))
    Case STOPED
        Call BASS_ChannelSetPosition(Mp3(1).chan, Mp3(1).file.SongStart)
        Call BASS_StreamPlay(BASS_FX_TempoGetResampledHandle(Mp3(1).chan), 0, BASS_STREAM_AUTOFREE)
        Call BASS_ChannelPause(BASS_FX_TempoGetResampledHandle(Mp3(1).chan))
        Call BASS_ChannelResume(BASS_FX_TempoGetResampledHandle(Mp3(1).chan))
    Case PAUSED
        Call BASS_ChannelResume(BASS_FX_TempoGetResampledHandle(Mp3(1).chan))
    End Select

    MixerFrm.CrossFader_ValueChanged
Mp3(1).Estado = PLAYING
Label2.Caption = "DECK A : PLAYING"
'update pitch view
Call pitchSlider_ValueChanged
MeReadyToStop = False
End Function





Private Sub analyzer_DragDrop(Source As Control, X As Single, Y As Single)
If Source.Name = "ListView2" Then
    IsDrag = True
    LoadMP3
    IsDrag = False
    Set Master.ListView2.DropHighlight = Nothing
End If
End Sub








Private Sub cmdResetBPM_Click()
If Mp3(1).Estado <> IDLE Then
    pitchSlider.value = 0
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DeckLoad_Click()
LoadMP3
End Sub

Private Sub DeckPlay_Click()
PlayMP3
End Sub



Private Sub Form_Load()
'Mp3(1).Estado = IDLE
'Me.Height = 2490
'PaintBackground Me, BgSrc
End Sub

Public Sub fx1_Click()
Mp3(1).fx1 = Not Mp3(1).fx1
If Mp3(1).fx1 = True Then
    efx1 True
    fx1.BackColor = RGB(0, 255, 0)
Else
    efx1 False
    fx1.BackColor = &HC0C0C0
End If
End Sub

Private Sub fx2_Click()
Mp3(1).fx2 = Not Mp3(1).fx2
'///////////// Flanger //////////////
    If Mp3(1).fx2 = True Then
        efx2 True
        fx2.BackColor = RGB(0, 255, 0)
    Else
        efx2 False
        fx2.BackColor = &HC0C0C0
    End If
End Sub
Private Sub Picture5_DragDrop(Source As Control, X As Single, Y As Single)
If Source.Name = "ListView2" Then
    IsDrag = True
    LoadMP3
    IsDrag = False
    Set Master.ListView2.DropHighlight = Nothing
End If
End Sub

Private Sub fx3_Click()
Mp3(1).fx3 = Not Mp3(1).fx3
'///////////// brake //////////////
If Mp3(1).fx3 = True Then
        fx3.BackColor = RGB(0, 255, 0)
        efx3 True
Else
        efx3 False
        fx3.BackColor = &HC0C0C0
End If
End Sub

Private Sub fx4_Click()
Mp3(1).fx4 = Not Mp3(1).fx4
If Mp3(1).fx4 = True Then
     efx4 True
    fx4.BackColor = RGB(0, 255, 0)
Else
    efx4 False
    fx4.BackColor = &HC0C0C0
End If
End Sub

Private Sub fx5_Click()
Mp3(1).fx5 = Not Mp3(1).fx5
'///////////// brake //////////////
If Mp3(1).fx5 = True Then
        efx5 True
        fx5.BackColor = RGB(0, 255, 0)
Else
        efx5 False
        fx5.BackColor = &HC0C0C0
End If
End Sub









Private Sub mnuExit_Click()
    End
End Sub

Private Sub pitchSlider_ValueChanged()
    Call BASS_FX_TempoSet(Mp3(1).chan, CLng(pitchSlider.value), -1, -100#)
    Mp3(1).newBPM = GetNewBPM1()
End Sub

Private Sub pos_MouseDown(Shift As Integer)
    Timer1.Enabled = False
End Sub

Private Sub pos_MouseUp(Shift As Integer)
    Call BASS_ChannelSetPosition(Mp3(1).chan, pos.value)
    Timer1.Enabled = True
End Sub

Private Sub StopDeck_Click()
StopMP3
End Sub

Public Sub Timer1_Timer()

If BASS_ChannelIsActive(Mp3(1).chan) = BASS_ACTIVE_PLAYING Then '
pos.value = BASS_ChannelGetPosition(Mp3(1).chan)
lblBPM.Caption = Mp3(1).newBPM

    '---------------- MP3 --------------------------
    If AutoDJ = True And Mp3(2).Estado = IDLE And pos.value >= LoadPoint Then
        Player2.LoadMP3
    End If
    
    If AutoDJ = True And MixOut = True And Player2.MeReadyToStop = True Then
        MixOut = False
        Player2.StopMP3
        Player2.UnloadMP3
        BASS_ChannelRemoveSync Mp3(1).chan, SNYC1a
    End If
    
    If AutoDJ = True And Mix = True Then
        Mix = False
        MeReadyToStop = True
        Player2.PlayMP3
        Call BASS_ChannelSlideAttributes(BASS_FX_TempoGetResampledHandle(Mp3(1).chan), -1, -2, -101, 1000)
        BASS_ChannelRemoveSync Mp3(1).chan, SNYC1
    
    End If
    
    If Master.enabSpectrum.Checked = True Then
    'spectrum analyzer BASS_DATA_FFT512
    BASS_ChannelGetData BASS_FX_TempoGetResampledHandle(Mp3(1).chan), SampleTemp(0), BASS_DATA_FFT512
    'display spectrum analyzer
    For i = 0 To 35
'        bitblt abuff.hdc, 4 * i, 65 - (Sqrt(SampleTemp(i)) * 140), 3, (Sqrt(SampleTemp(i)) * 140) + 1, gph.hdc, 0, 0, vbSrcCopy
        bitblt abuff.hdc, i * 5, 65 - (Sqrt(SampleTemp(i)) * 140), 4, (Sqrt(SampleTemp(i)) * 140) + 1, gph.hdc, 0, 65 - (Sqrt(SampleTemp(i)) * 140), vbSrcCopy
        
    Next i
    bitblt analyzer.hdc, 0, 0, 170, 100, abuff.hdc, 0, 0, vbSrcCopy
    abuff.Cls
    End If
    
Else
    Timer1.Enabled = False
End If
End Sub

Public Function Sqrt(ByVal num As Double) As Double
    Sqrt = num ^ 0.5
End Function


'effects
Private Sub efx1(state As Boolean)
Dim rv2 As BASS_FX_DSPREVERB
If state = True Then
    Call BASS_FX_DSP_Set(Mp3(1).chan, BASS_FX_DSPFX_REVERB, 1)
    Call BASS_FX_DSP_GetParameters(Mp3(1).chan, BASS_FX_DSPFX_REVERB, rv2)
        rv2.fLevel = 0.5
        rv2.lDelay = 3000
    Call BASS_FX_DSP_SetParameters(Mp3(1).chan, BASS_FX_DSPFX_REVERB, rv2)
Else
    Call BASS_FX_DSP_Remove(Mp3(1).chan, BASS_FX_DSPFX_REVERB)
End If
End Sub

Private Sub efx2(state As Boolean)
Dim fl2 As BASS_FX_DSPFLANGER2
    If state = True Then
        Call BASS_FX_DSP_Set(Mp3(1).chan, BASS_FX_DSPFX_FLANGER2, 1)
        Call BASS_FX_DSP_GetParameters(Mp3(1).chan, BASS_FX_DSPFX_FLANGER2, fl2)
        Call BASS_ChannelGetAttributes(Mp3(1).chan, fl2.lFreq, 0, 0)
               fl2.fDelay = 210 / 100
               fl2.fBPM = 120
               fl2.fWetDry = 2
        Call BASS_FX_DSP_SetParameters(Mp3(1).chan, BASS_FX_DSPFX_FLANGER2, fl2)
    Else
        Call BASS_FX_DSP_Remove(Mp3(1).chan, BASS_FX_DSPFX_FLANGER2)
    End If
End Sub

Private Sub efx3(state As Boolean)
Dim fl2 As BASS_FX_DSPECHO
    If state = True Then
        Call BASS_FX_DSP_Set(Mp3(1).chan, BASS_FX_DSPFX_ECHO, 1)
        Call BASS_FX_DSP_GetParameters(Mp3(1).chan, BASS_FX_DSPFX_ECHO, fl2)
            fl2.fLevel = 0.5
            fl2.lDelay = 8000
               
        Call BASS_FX_DSP_SetParameters(Mp3(1).chan, BASS_FX_DSPFX_ECHO, fl2)
    Else
        Call BASS_FX_DSP_Remove(Mp3(1).chan, BASS_FX_DSPFX_ECHO)
    End If
End Sub
Private Sub efx4(state As Boolean)
If state = True Then
    Call BASS_FX_TempoSet(Mp3(1).chan, -100#, -1, 8)
Else
    Call BASS_FX_TempoSet(Mp3(1).chan, -100#, -1, 0)
End If
End Sub

Private Sub efx5(state As Boolean)
Dim damp As BASS_FX_DSPDAMP
If state = True Then
        Call BASS_FX_DSP_Set(Mp3(1).chan, BASS_FX_DSPFX_DAMP, 1)
        Call BASS_FX_DSP_GetParameters(Mp3(1).chan, BASS_FX_DSPFX_DAMP, damp)
               damp.fGain = 1
               damp.fRate = 0.02
               damp.lDelay = 1000
               damp.lQuiet = 300
               damp.lTarget = 30000
        Call BASS_FX_DSP_SetParameters(Mp3(1).chan, BASS_FX_DSPFX_DAMP, damp)
Else
        Call BASS_FX_DSP_Remove(Mp3(1).chan, BASS_FX_DSPFX_DAMP)
End If
End Sub


Private Sub PaintBackground(dfrm As Form, BgSrc As PictureBox)
Dim X As Integer
Dim Y As Integer
For Y = 0 To dfrm.ScaleHeight Step BgSrc.ScaleHeight
    For X = 0 To dfrm.ScaleWidth Step BgSrc.ScaleWidth
         bitblt dfrm.hdc, X, Y, BgSrc.ScaleWidth, BgSrc.ScaleHeight, BgSrc.hdc, 0, 0, vbSrcCopy
         DoEvents
    Next X
Next Y
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
    Case Is = "play"
        'PlayNow
    Case Is = "pause"
        Select Case Mp3(1).Estado
        Case PLAYING
            Player1.PauseMP3
        End Select
        Select Case Mp3(2).Estado
        Case PLAYING
            Player2.PauseMP3
        End Select
    Case Is = "stop"
        Select Case Mp3(1).Estado
        Case PLAYING
            Player1.StopMP3
        End Select
        Select Case Mp3(2).Estado
        Case PLAYING
            Player2.StopMP3
        End Select
    Case Is = "next"
        MixerFrm.MixNow

    Case Is = "autodj"
        AutoDJ = Not AutoDJ
        If AutoDJ = True Then
            MixerFrm.lblAutoDJ.ForeColor = &HFFFF00
        Else
            MixerFrm.lblAutoDJ.ForeColor = &H808000
        End If
        
    Case Is = "equalizer"
        EqualizerFrm.Show
        
    Case Is = "mic"
        MsgBox "This function is disabled on this version", vbInformation Or vbOKOnly, "Information"
    Case Is = "audio"
        If FileExist("C:\Windows\sndvol32.exe") = True Then
            Shell "sndvol32.exe", vbNormalFocus
        End If
    Case Is = "browse"
        'mnuBrowser_Click
    Case Is = "mixer"
        'MixerClick
 
 Case Is = "save"
    'SaveCurrentList
 Case Is = "load"
    'LoadUserPlaylist
 Case Is = "wizard"
     PLaylistGeneratorFrm.Show
 Case Is = "library"
     SSTab1.Tab = 0
 Case Is = "radio"
     SSTab1.Tab = 1
 Case Is = "history"
     SSTab1.Tab = 2
 End Select

End Sub
