VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "MGC DJ 2000"
   ClientHeight    =   8340
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   11595
   Icon            =   "MGC-DJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "MGC-DJ.frx":030A
   ScaleHeight     =   8340
   ScaleMode       =   0  'User
   ScaleWidth      =   11599
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   7200
      TabIndex        =   4
      ToolTipText     =   "Search Music Box B / Press ""F5"" to play song selected"
      Top             =   6840
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "Search Music Box A / Press ""F5 "" to play song selected"
      Top             =   6480
      Width           =   3495
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7920
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tupdate 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1440
      Top             =   0
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00000000&
      Caption         =   "DJ Mixer"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   98
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.DriveListBox Drive3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   270
      Left            =   360
      TabIndex        =   21
      ToolTipText     =   "Music Drive Box Special"
      Top             =   720
      Width           =   2535
   End
   Begin VB.DirListBox Dir3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   450
      Left            =   360
      TabIndex        =   22
      ToolTipText     =   "Music Directory Box Special"
      Top             =   960
      Width           =   2535
   End
   Begin VB.FileListBox File3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1140
      Hidden          =   -1  'True
      Left            =   360
      System          =   -1  'True
      TabIndex        =   23
      ToolTipText     =   "Music Box Special / Make Double Click or press F5 or ENTER"
      Top             =   1400
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   360
      MaxLength       =   20
      TabIndex        =   24
      ToolTipText     =   "Search Music Box Special / Press ""F5"" to play song selected"
      Top             =   2520
      Width           =   2535
   End
   Begin ComctlLib.Slider Slider3 
      Height          =   255
      Left            =   7185
      TabIndex        =   8
      ToolTipText     =   "Track Position B"
      Top             =   7080
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      MouseIcon       =   "MGC-DJ.frx":92A6
      LargeChange     =   30
      SmallChange     =   5
      SelectRange     =   -1  'True
      SelLength       =   10
      TickStyle       =   3
      TickFrequency   =   10
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   255
      Left            =   345
      TabIndex        =   6
      ToolTipText     =   "Track Position A"
      Top             =   6720
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      MouseIcon       =   "MGC-DJ.frx":95C0
      LargeChange     =   30
      SmallChange     =   5
      SelectRange     =   -1  'True
      SelLength       =   10
      TickStyle       =   3
      TickFrequency   =   10
   End
   Begin ComctlLib.Slider Slider7 
      Height          =   255
      Left            =   345
      TabIndex        =   26
      ToolTipText     =   "Track Position Special"
      Top             =   2760
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      MouseIcon       =   "MGC-DJ.frx":98DA
      LargeChange     =   30
      SmallChange     =   5
      SelectRange     =   -1  'True
      SelLength       =   10
      TickStyle       =   3
      TickFrequency   =   15
   End
   Begin VB.CommandButton Command9 
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   3.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   170
      Left            =   3960
      TabIndex        =   94
      ToolTipText     =   "Clear the Playlist A"
      Top             =   4560
      Width           =   400
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   3.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   170
      Left            =   3960
      TabIndex        =   88
      ToolTipText     =   "Save a MGC DJ Playlist"
      Top             =   4440
      Width           =   400
   End
   Begin VB.CommandButton Command5 
      Caption         =   "LOAD"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   3.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   170
      Left            =   3960
      TabIndex        =   89
      ToolTipText     =   "Load a MGC DJ Playlist"
      Top             =   4320
      Width           =   400
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   3.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   170
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Add a Music File to Playlist A"
      Top             =   4200
      Width           =   400
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00000000&
      Caption         =   "Special Mix On"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   7440
      TabIndex        =   97
      ToolTipText     =   "Special Mix Effect - It works with Mixer Effect On and Mixer Effect Scroll"
      Top             =   7680
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   3.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   170
      Left            =   6720
      TabIndex        =   93
      ToolTipText     =   "Clear the Playlist B"
      Top             =   4560
      Width           =   400
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   3.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   170
      Left            =   6720
      TabIndex        =   90
      ToolTipText     =   "Save a MGC DJ Playlist"
      Top             =   4440
      Width           =   400
   End
   Begin VB.CommandButton Command7 
      Caption         =   "LOAD"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   3.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   170
      Left            =   6720
      TabIndex        =   91
      ToolTipText     =   "Load a MGC DJ Playlist"
      Top             =   4320
      Width           =   400
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   3.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   170
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   92
      ToolTipText     =   "Add a Music File to Playlist B"
      Top             =   4200
      Width           =   400
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   360
      TabIndex        =   86
      ToolTipText     =   "Music Box A / Make Double Click or press F5 or ENTER"
      Top             =   3840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4683
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   65280
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Music Files A"
         Object.Width           =   5610
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Full Path"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox tt 
      Height          =   285
      Left            =   4680
      TabIndex        =   85
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox text25 
      Height          =   285
      Left            =   5160
      TabIndex        =   84
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   5880
      TabIndex        =   83
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   6360
      TabIndex        =   82
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1115
      Left            =   50000
      Top             =   0
   End
   Begin ComctlLib.Slider Slider8 
      Height          =   615
      Left            =   3000
      TabIndex        =   28
      ToolTipText     =   "Volume Special"
      Top             =   2280
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1085
      _Version        =   327682
      BorderStyle     =   1
      Orientation     =   1
      LargeChange     =   500
      SmallChange     =   100
      Min             =   -3000
      Max             =   0
      SelectRange     =   -1  'True
      SelStart        =   -3000
      SelLength       =   3000
      TickFrequency   =   500
   End
   Begin ComctlLib.Slider Slider11 
      Height          =   975
      Left            =   3480
      TabIndex        =   59
      ToolTipText     =   "Master Volume"
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1720
      _Version        =   327682
      BorderStyle     =   1
      Orientation     =   1
      LargeChange     =   500
      SmallChange     =   100
      Min             =   -3000
      Max             =   0
      SelectRange     =   -1  'True
      SelStart        =   -3000
      SelLength       =   3000
      TickStyle       =   3
   End
   Begin ComctlLib.Slider Vol1 
      Height          =   615
      Left            =   3960
      TabIndex        =   10
      ToolTipText     =   "Volume A"
      Top             =   6360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1085
      _Version        =   327682
      BorderStyle     =   1
      Orientation     =   1
      LargeChange     =   500
      SmallChange     =   100
      Min             =   -3000
      Max             =   0
      SelectRange     =   -1  'True
      SelStart        =   -3000
      SelLength       =   3000
      TickFrequency   =   500
   End
   Begin ComctlLib.Slider Vol2 
      Height          =   615
      Left            =   6840
      TabIndex        =   11
      ToolTipText     =   "Volume B"
      Top             =   6720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1085
      _Version        =   327682
      BorderStyle     =   1
      Orientation     =   1
      LargeChange     =   500
      SmallChange     =   100
      Min             =   -3000
      Max             =   0
      SelectRange     =   -1  'True
      SelStart        =   -3000
      SelLength       =   3000
      TickFrequency   =   500
   End
   Begin VB.Timer Timerxx 
      Interval        =   200
      Left            =   2160
      Top             =   0
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   350
      Left            =   5100
      TabIndex        =   1
      ToolTipText     =   "Mixer Effect Scroll"
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      LargeChange     =   500
      SmallChange     =   100
      Min             =   -3000
      Max             =   0
      SelectRange     =   -1  'True
      SelStart        =   -3000
      TickFrequency   =   250
   End
   Begin ComctlLib.Slider Slider6 
      Height          =   345
      Left            =   5160
      TabIndex        =   18
      ToolTipText     =   "Balance Effect Scroll"
      Top             =   5040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      LargeChange     =   1000
      SmallChange     =   500
      Min             =   -10000
      Max             =   10000
      SelectRange     =   -1  'True
      SelStart        =   -10000
      SelLength       =   10000
      TickFrequency   =   1500
      Value           =   -3000
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00FFFFFF&
      Caption         =   "> > >"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   170
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Auto Mix Effect Right / Press ""Left Alt"" Key"
      Top             =   5570
      Width           =   615
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      MaxLength       =   2
      TabIndex        =   55
      Text            =   "50"
      ToolTipText     =   "Select the Mix Speed (1-99)"
      Top             =   5550
      Width           =   250
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2400
      Top             =   7920
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "< < <"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   170
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Auto Mix Effect Left / Press ""Left Ctrl"" Key"
      Top             =   5570
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Balance Effect Off"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4800
      MaskColor       =   &H00FF0000&
      TabIndex        =   53
      ToolTipText     =   "Allows you to make a special Balance Effect. It work with the Balance Effect Scroll"
      Top             =   2640
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Mixer  Effect On"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6240
      MaskColor       =   &H00FF0000&
      TabIndex        =   52
      ToolTipText     =   "Allows you to Mix music between Playlist A and B"
      Top             =   2520
      UseMaskColor    =   -1  'True
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   2.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   90
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Exit"
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   90
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Minimize"
      Top             =   120
      Width           =   255
   End
   Begin ComctlLib.Slider Slider10 
      Height          =   105
      Left            =   8880
      TabIndex        =   32
      ToolTipText     =   "Dj´s Effect Volume"
      Top             =   3000
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   185
      _Version        =   327682
      LargeChange     =   100
      SmallChange     =   50
      Min             =   -3000
      Max             =   0
      SelectRange     =   -1  'True
      SelStart        =   -3000
      SelLength       =   3000
      TickStyle       =   3
   End
   Begin VB.Timer contador 
      Interval        =   9
      Left            =   3360
      Top             =   -600
   End
   Begin VB.Timer Timer0 
      Interval        =   19
      Left            =   3000
      Top             =   -600
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   600
      ScaleHeight     =   375
      ScaleWidth      =   3255
      TabIndex        =   25
      Top             =   3240
      Width           =   3255
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   42
         Text            =   "0"
         ToolTipText     =   "Minute Effect Special / Select the desired minute time to begin the song and press ""Play"""
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   44
         Text            =   "0"
         ToolTipText     =   "Second Effect Special / Select the desired second time to begin the song and press ""Play"""
         Top             =   120
         Width           =   250
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1830
         TabIndex        =   43
         ToolTipText     =   "Double Click: Reset Min Sec Effect Special"
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "0:00 / 0:00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         ToolTipText     =   "Track Duration Special"
         Top             =   120
         Width           =   975
      End
      Begin VB.Image Image6 
         Height          =   375
         Left            =   960
         Picture         =   "MGC-DJ.frx":9BF4
         ToolTipText     =   "Pause Special"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   375
         Left            =   480
         Picture         =   "MGC-DJ.frx":9E27
         ToolTipText     =   "Stop Special"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Image4 
         Height          =   375
         Left            =   0
         Picture         =   "MGC-DJ.frx":A02D
         ToolTipText     =   "Play Special"
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CheckBox repeat 
      BackColor       =   &H80000007&
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      ToolTipText     =   "Continue Playing Next Song on playlists"
      Top             =   7560
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin ComctlLib.Slider Slider5 
      Height          =   135
      Left            =   6840
      TabIndex        =   17
      Top             =   7400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   238
      _Version        =   327682
      LargeChange     =   1000
      SmallChange     =   500
      Min             =   -10000
      Max             =   10000
      SelectRange     =   -1  'True
      SelStart        =   -10000
      SelLength       =   10000
   End
   Begin ComctlLib.Slider Slider4 
      Height          =   135
      Left            =   3240
      TabIndex        =   16
      Top             =   7320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   238
      _Version        =   327682
      LargeChange     =   1000
      SmallChange     =   500
      Min             =   -10000
      Max             =   10000
      SelectRange     =   -1  'True
      SelStart        =   -10000
      SelLength       =   10000
   End
   Begin VB.Timer Timer5 
      Interval        =   24
      Left            =   480
      Top             =   -600
   End
   Begin VB.Timer Timer4 
      Interval        =   18
      Left            =   -240
      Top             =   -600
   End
   Begin VB.Timer Timer3 
      Interval        =   400
      Left            =   120
      Top             =   -600
   End
   Begin VB.Timer Tmr2Time 
      Interval        =   2
      Left            =   1200
      Top             =   -600
   End
   Begin VB.Timer Timet1 
      Interval        =   5
      Left            =   1920
      Top             =   -600
   End
   Begin VB.Timer TmrTime 
      Interval        =   10
      Left            =   840
      Top             =   -600
   End
   Begin VB.Timer Mousetmr 
      Interval        =   15
      Left            =   2640
      Top             =   -600
   End
   Begin VB.Timer Timer2 
      Interval        =   31
      Left            =   1560
      Top             =   -600
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   2280
      Top             =   -600
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   2640
      TabIndex        =   5
      Top             =   7320
      Width           =   2640
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   37
         Text            =   "0"
         ToolTipText     =   "Second Effect A / Select the desired second time to begin the song and press ""Play"""
         Top             =   0
         Width           =   250
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   36
         Text            =   "0"
         ToolTipText     =   "Minute Effect A / Select the desired minute time to begin the song and press ""Play"""
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   2070
         TabIndex        =   38
         ToolTipText     =   "Double Click: Reset Min Sec Effect A"
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "0:00 / 0:00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         ToolTipText     =   "Track Duration A"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image B_Pause 
         Height          =   375
         Left            =   1080
         Picture         =   "MGC-DJ.frx":A213
         ToolTipText     =   "Pause A"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image B_Stop 
         Height          =   375
         Left            =   600
         Picture         =   "MGC-DJ.frx":A446
         ToolTipText     =   "Stop A"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image B_Play 
         Height          =   375
         Left            =   120
         Picture         =   "MGC-DJ.frx":A64C
         ToolTipText     =   "Play A"
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   8760
      ScaleHeight     =   495
      ScaleWidth      =   2565
      TabIndex        =   7
      Top             =   7560
      Width           =   2566
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   39
         Text            =   "0"
         ToolTipText     =   "Minute Effect B / Select the desired minute time to begin the song and press ""Play"""
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   41
         Text            =   "0"
         ToolTipText     =   "Second Effect B / Select the desired second time to begin the song and press ""Play"""
         Top             =   0
         Width           =   250
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2070
         TabIndex        =   40
         ToolTipText     =   "Double Click: Reset Min Sec Effect B"
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "0:00 / 0:00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1500
         TabIndex        =   13
         ToolTipText     =   "Track Duration B"
         Top             =   210
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   120
         Picture         =   "MGC-DJ.frx":A832
         ToolTipText     =   "Play B"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   600
         Picture         =   "MGC-DJ.frx":AA18
         ToolTipText     =   "Stop B"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   1080
         Picture         =   "MGC-DJ.frx":AC1E
         ToolTipText     =   "Pause B"
         Top             =   0
         Width           =   375
      End
   End
   Begin ComctlLib.Slider Slider9 
      Height          =   135
      Left            =   3000
      TabIndex        =   29
      Top             =   3000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   238
      _Version        =   327682
      LargeChange     =   1000
      SmallChange     =   500
      Min             =   -10000
      Max             =   10000
      SelectRange     =   -1  'True
      SelStart        =   -10000
      SelLength       =   10000
   End
   Begin MSComctlLib.Slider Sliderxx 
      Height          =   255
      Left            =   7905
      TabIndex        =   58
      ToolTipText     =   "You can choose a beginning for playing the recording"
      Top             =   3480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   500
      SmallChange     =   100
      SelectRange     =   -1  'True
      TickStyle       =   3
   End
   Begin MSComDlg.CommonDialog CommonDialogxx 
      Left            =   2640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   " "
      Orientation     =   2
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2655
      Left            =   7200
      TabIndex        =   87
      ToolTipText     =   "Music Box B / Make Double Click or press F5 or ENTER"
      Top             =   4200
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4683
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   65535
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Music Files B"
         Object.Width           =   6439
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Full Path"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Image Image32 
      Height          =   195
      Left            =   9600
      Picture         =   "MGC-DJ.frx":AE51
      Stretch         =   -1  'True
      ToolTipText     =   "Stop Dj´s Effect"
      Top             =   2920
      Width           =   135
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   10440
      TabIndex        =   107
      ToolTipText     =   "Setup the Record Options"
      Top             =   3285
      Width           =   735
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   9840
      TabIndex        =   106
      ToolTipText     =   "Save the data recorded on a wav file"
      Top             =   3285
      Width           =   615
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   9360
      TabIndex        =   105
      ToolTipText     =   "Play the data recorded"
      Top             =   3280
      Width           =   495
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   9000
      TabIndex        =   104
      ToolTipText     =   "Stop the Record or Play"
      Top             =   3280
      Width           =   375
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Record"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   8447
      TabIndex        =   103
      ToolTipText     =   "Start the Record Process"
      Top             =   3280
      Width           =   495
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   7920
      TabIndex        =   102
      ToolTipText     =   "Clean all data recorded"
      Top             =   3280
      Width           =   495
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mixer Effect On"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   6360
      TabIndex        =   101
      ToolTipText     =   "Allows you to Mix music between Playlist A and B"
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Effect Off"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4320
      TabIndex        =   100
      ToolTipText     =   "Allows you to make a special Balance Effect. It work with the Balance Effect Scroll"
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mix Mode 1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5760
      TabIndex        =   99
      ToolTipText     =   "Select Mix Mode - It works with Mixer Effect On and Mixer Effect Scroll"
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label playl1 
      Height          =   135
      Left            =   8040
      TabIndex        =   96
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label playl2 
      Height          =   255
      Left            =   7080
      TabIndex        =   95
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Labeldj 
      Height          =   135
      Left            =   3720
      TabIndex        =   80
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OPTIONS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   7440
      TabIndex        =   79
      ToolTipText     =   "Options Menu (Make Click on Right Mouse Bottom)"
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label StatisticsLabel 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   7920
      TabIndex        =   57
      ToolTipText     =   "Information about the recording"
      Top             =   3720
      Width           =   3255
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer2 
      Height          =   30
      Left            =   -480
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   30
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   30
      Left            =   -480
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   30
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   -1  'True
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer3 
      Height          =   495
      Left            =   -360
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer4 
      Height          =   495
      Left            =   -360
      TabIndex        =   31
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Label MGC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4320
      TabIndex        =   51
      ToolTipText     =   "Dj Text Section"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   360
      TabIndex        =   50
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label16 
      Caption         =   "                                                                                                    "
      Height          =   255
      Left            =   1200
      TabIndex        =   49
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label15 
      Height          =   255
      Left            =   3360
      TabIndex        =   48
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label14 
      Height          =   135
      Left            =   8000
      TabIndex        =   45
      Top             =   8520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label10 
      Height          =   135
      Left            =   7080
      TabIndex        =   35
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label cp 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "C.P. On"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4515
      TabIndex        =   34
      ToolTipText     =   "Continue Playing Next Song on playlists"
      Top             =   7150
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "DJ´s Effects"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9600
      TabIndex        =   33
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   6360
      TabIndex        =   19
      Top             =   -480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7920
      TabIndex        =   15
      Top             =   7320
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      ToolTipText     =   "Time"
      Top             =   1850
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   6960
      Width           =   3495
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   60
      ToolTipText     =   "Press F1"
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   61
      ToolTipText     =   "Press F2"
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   62
      ToolTipText     =   "Press F3"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   63
      ToolTipText     =   "Press F4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   64
      ToolTipText     =   "Press F6"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   65
      ToolTipText     =   "Press F7"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   66
      ToolTipText     =   "Press F8"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   67
      ToolTipText     =   "Press F9"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   68
      ToolTipText     =   "Press F10"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10080
      TabIndex        =   78
      ToolTipText     =   "Press F11"
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10080
      TabIndex        =   69
      ToolTipText     =   "Press F12"
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10080
      TabIndex        =   70
      ToolTipText     =   "Press Shift + F1"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10080
      TabIndex        =   71
      ToolTipText     =   "Press Shift + F2"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10080
      TabIndex        =   72
      ToolTipText     =   "Press Shift + F3"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10080
      TabIndex        =   73
      ToolTipText     =   "Press Shift + F4"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10080
      TabIndex        =   74
      ToolTipText     =   "Press Shift + F6"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10080
      TabIndex        =   75
      ToolTipText     =   "Press Shift + F7"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10080
      TabIndex        =   76
      ToolTipText     =   "Press Shift + F8"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10080
      TabIndex        =   77
      ToolTipText     =   "Press Shift + F9"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Image Image7 
      Height          =   195
      Left            =   8880
      Picture         =   "MGC-DJ.frx":B07A
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image Image8 
      Height          =   195
      Left            =   8880
      Picture         =   "MGC-DJ.frx":B2A3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Image Image9 
      Height          =   195
      Left            =   8880
      Picture         =   "MGC-DJ.frx":B4CC
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image Image10 
      Height          =   195
      Left            =   8880
      Picture         =   "MGC-DJ.frx":B6F5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image Image11 
      Height          =   195
      Left            =   8880
      Picture         =   "MGC-DJ.frx":B91E
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image Image12 
      Height          =   195
      Left            =   8880
      Picture         =   "MGC-DJ.frx":BB47
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Image Image13 
      Height          =   195
      Left            =   8880
      Picture         =   "MGC-DJ.frx":BD70
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Image Image14 
      Height          =   195
      Left            =   8880
      Picture         =   "MGC-DJ.frx":BF99
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Image Image15 
      Height          =   195
      Left            =   8880
      Picture         =   "MGC-DJ.frx":C1C2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Image Image16 
      Height          =   195
      Left            =   10080
      Picture         =   "MGC-DJ.frx":C3EB
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image Image17 
      Height          =   195
      Left            =   10080
      Picture         =   "MGC-DJ.frx":C614
      Top             =   960
      Width           =   1095
   End
   Begin VB.Image Image18 
      Height          =   195
      Left            =   10080
      Picture         =   "MGC-DJ.frx":C83D
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image Image19 
      Height          =   195
      Left            =   10080
      Picture         =   "MGC-DJ.frx":CA66
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image Image20 
      Height          =   195
      Left            =   10080
      Picture         =   "MGC-DJ.frx":CC8F
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image Image21 
      Height          =   195
      Left            =   10080
      Picture         =   "MGC-DJ.frx":CEB8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Image Image22 
      Height          =   195
      Left            =   10080
      Picture         =   "MGC-DJ.frx":D0E1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Image Image23 
      Height          =   195
      Left            =   10080
      Picture         =   "MGC-DJ.frx":D30A
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Image Image24 
      Height          =   195
      Left            =   10080
      Picture         =   "MGC-DJ.frx":D533
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Image Image25 
      Height          =   195
      Left            =   10080
      Picture         =   "MGC-DJ.frx":D75C
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Image Image26 
      Height          =   230
      Left            =   7920
      Picture         =   "MGC-DJ.frx":D985
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   495
   End
   Begin VB.Image Image27 
      Height          =   225
      Left            =   8400
      Picture         =   "MGC-DJ.frx":DBAE
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image Image28 
      Height          =   225
      Left            =   9000
      Picture         =   "MGC-DJ.frx":DDD7
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image Image29 
      Height          =   225
      Left            =   9360
      Picture         =   "MGC-DJ.frx":E000
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   495
   End
   Begin VB.Image Image30 
      Height          =   225
      Left            =   9840
      Picture         =   "MGC-DJ.frx":E229
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image Image31 
      Height          =   225
      Left            =   10440
      Picture         =   "MGC-DJ.frx":E452
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim mleft As Integer
Dim dirc As String
Dim indrag As Boolean 'Indicador de operación de arrastrar y colocar
Dim nodX As Object
Dim tempo2 As Integer
Dim listado As Integer
Dim listado1 As Integer
Dim listado2 As Integer
Dim mousedown As Integer
Dim MouseDown1 As Integer
Dim durac As Integer
Dim durac2 As Integer
Dim listTemp2() As String
Dim Direction As String
Dim duracion As Integer
Const NO_BUTTON = 0
Const WM_NCLBUTTONDOWN = &HA1
Const MouseL = 1
Const MouseR = 2
Const Key_Up = &H26
Const Key_Down = &H28
Const Key_Enter = &HD
Const Key_F5 = &H74
Const Key_RePag = &H21
Const Key_AvPag = &H22
Private Sub B_Pause_Click()
On Error GoTo bye
MediaPlayer1.Pause
bye:
End Sub
Function StripItem(startStrg As String, parser As String) As String
'this takes a string separated by the chr passed in Parser,
'splits off 1 item, and shortens startStrg so that the next
'item is ready for removal.

   Dim C As Integer
   Dim item As String
   
   C = 1
   
   Do
   
      If Mid(startStrg, C, 1) = parser Then
      
         item = Mid(startStrg, 1, C - 1)
         startStrg = Mid(startStrg, C + 1, Len(startStrg))
         StripItem = item
         Exit Function
      End If
      
      C = C + 1
   
   Loop

End Function

Private Sub B_Play_Click()
On Error GoTo bye
MediaPlayer1.Play
If Text4.Text = "0" And Text5.Text = "0" Then Exit Sub
MediaPlayer1.CurrentPosition = Text4.Text * 60 + Text5.Text
bye:
End Sub

Private Sub B_Stop_Click()
MediaPlayer1.Stop
MediaPlayer1.CurrentPosition = 0
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Check2.ForeColor = &HFF00&
If Label14.Caption = "Sp" Then
Check2.Caption = "Efecto Balance On"
Else
Check2.Caption = "Balance Effect On"
End If
Slider6.SetFocus
Else
Check2.ForeColor = &HFF&
If Label14.Caption = "Sp" Then
Check2.Caption = "Efecto Balance Off"
Else
Check2.Caption = "Balance Effect Off"
End If
Slider4.Value = MediaPlayer1.Balance
Slider5.Value = MediaPlayer2.Balance
End If
End Sub





Private Sub Check1_Click()
If Check1.Value = 1 Then
Check1.ForeColor = &HFF00&
If Label14.Caption = "Sp" Then
Check1.Caption = "Efecto  Mixer On"
Else
Check1.Caption = "Mixer  Effect On"
End If
Vol1.Enabled = False
Vol2.Enabled = False
Slider1.SetFocus
Else
Check1.ForeColor = &HFF&
If Label14.Caption = "Sp" Then
Check1.Caption = "Efecto  Mixer Off"
Else
Check1.Caption = "Mixer  Effect Off"
End If
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Vol1.Enabled = True
Vol2.Enabled = True
End If
End Sub



Private Sub Check3_Click()
If Check3.Value = 1 Then Check3.Caption = "Special Mix On": Check3.ForeColor = &HFF00&: Exit Sub
Check3.Caption = "Special Mix Off"
Check3.ForeColor = &HFF&
End Sub




























Private Sub Command2_Click()
  'working variables
      Dim C As Integer
   Dim q As Integer
   Dim sFile As String
   Dim startStrg As String
   Dim tmp As String
  
  'dim an array to hold the files selected
   Dim FileArray() As String
   ListView1.Sorted = False
   'Timer1.Enabled = False
   'txtTime.Text = ""
   'txtCount.Text = ""
   'Check1.Value = 0
  'set the max buffer large enough to retrieve multiple files
   CommonDialog1.DialogTitle = "Adding song to your playlist." & "  Hold the ctrl key to multi-select."
   CommonDialog1.MaxFileSize = 16384
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "ALL Formats Suported (*.mp3;*.mpg;*.mpeg;*.m1v;*.mp2;*.mpa;*.mp3;*.wav;*.avi;*.asf;*.asx;*.rmi;*.wma;*.wax;*.mid;*.rmi;*.qt;*.aif;*.aifc;*.aiff;*.mov;*.au;*.snd)|*.mp3;*.mpg;*.mpeg;*.m1v;*.mp2;*.mpa;*.mp3;*.wav;*.avi;*.asf;*.asx;*.rmi;*.wma;*.wax;*.mid;*.rmi;*.qt;*.aif;*.aifc;*.aiff;*.mov;*.au;*.snd|MPEG Files (*.mp3;*.mpg;*.mpeg;*.m1v;*.mp2;*.mpa;*.mp3)|*.mp3;*.mpg;*.mpeg;*.m1v;*.mp2;*.mpa;*.mp3|WAV Files (*.wav;*.avi;*.asf;*.asx;*.rmi;*.wma;*.wax)|*.wav;*.avi;*.asf;*.asx;*.rmi;*.wma;*.wax|MIDI Files (*.mid;*.rmi)|*.mid;*.rmi|AIFF Files (*.qt;*.aif;*.aifc;*.aiff;*.mov)|*.qt;*.aif;*.aifc;*.aiff;*.mov|UNIX Files (*.au;*.snd)|*.au;*.snd|ALL Files (*.*)|*.*"
   CommonDialog1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
   CommonDialog1.ShowOpen     ' = 1
  If CommonDialog1.FileName = "" Then Exit Sub
      'assign the string returned from the
  'common dialog to startStrg for further processing.
  'Note that two "nulls" are appended to the
  'end of the string.  This is for use in the StripItem
  'routine below.
   startStrg = CommonDialog1.FileName & Chr(0) & Chr(0)
  
  'Extract each returned filename.
  'If only 1 file was selected, then the string
  'contains the fully-qualified path to the file.
   
  'If more than 1 string file was selected, the
  'string contains the path as the first item,
  'and the FileArray as the rest of the string,
  'each separated by a space.
   For C = 1 To Len(CommonDialog1.FileName)
      
     'extract 1 item from the string
      sFile = StripItem(startStrg, Chr(0))
      
     'if nothing's there, we're done
      If sFile = "" Then Exit For
      
        'redim the filename array
        'to add the new file. FileArray(0) is either the
        'path (if more than 1 file selected), or the
        'fully qualified filename (if only 1 file selected).
         ReDim Preserve FileArray(0 To q)
         FileArray(q) = LCase(sFile)
         
        'increment y by 1 for the next pass
         q = q + 1
      
      Next
      
     'display the results
      Text11.Text = ""
      Text12.Text = ""
      'List1.Clear
      
     'starting with 0, and ending with y-1 (because
     'the above will add 1 more 'y'
     'than there is actual array members)
      For C = 0 To q - 1
      
        'if its the first item, display it in Text1,
        'otherwise display the files selected in Text2.
         If C = 0 Then
               Text11.Text = FileArray(C)
         Else
         tmp = tmp & Text11.Text & "\" & FileArray(C) & vbCrLf
                  If Len(Text11.Text) = 3 Then GoTo falta
         For i = 1 To ListView1.ListItems.count
         If ListView1.ListItems(i).ListSubItems(1).Text = Text11.Text & "\" & FileArray(C) Then GoTo sit
                  Next i

'List1.AddItem Text11.Text & "\" & FileArray(c)
'List2.AddItem FileArray(c)


ListView1.ListItems.Add , , FileArray(C)
ListView1.ListItems(ListView1.ListItems.count).ListSubItems.Add , , Text11.Text & "\" & FileArray(C)

GoTo sit

falta:
For i = 1 To ListView1.ListItems.count
         If ListView1.ListItems(i).ListSubItems(1).Text = Text11.Text & FileArray(C) Then GoTo sit
                  Next i

ListView1.ListItems.Add , , FileArray(C)
ListView1.ListItems(ListView1.ListItems.count).ListSubItems.Add , , Text11.Text & FileArray(C)

'List1.AddItem Text11.Text & FileArray(c)
'List2.AddItem FileArray(c)
sit:
         End If
      Next
      Text12.Text = tmp
       If Text12.Text = "" Then
       For X = 1 To ListView1.ListItems.count
         If ListView1.ListItems(X).ListSubItems(1).Text = Text11.Text Then GoTo tre
                  Next X
       'List1.AddItem Text11.Text
       
        Dim s As String
    Dim Delimiter As Integer
    Dim e As Integer
    
    'Strip off drive letter
    s = Text11.Text

    For e = Len(s) To 0 Step -1
        If Mid(s, e, 1) = "\" Then
            Delimiter = Len(s) - e
            Exit For
        End If
    Next e

    s = Right(s, Delimiter)
        
       
       ListView1.ListItems.Add , , s
ListView1.ListItems(ListView1.ListItems.count).ListSubItems.Add , , Text11.Text

End If
tre:
    
       'If Text10.Text = "" Then Check1.Value = 1
       'txtCount.Text = List1.ListCount
       'If txtCount.Text = "1" Then
       'txtTime.Text = "0"
       'Timer1.Enabled = False
       'ElseIf txtCount.Text > "1" Then
       'txtTime.Text = "0"
       'Timer1.Enabled = True
       'End If
'       List1.ListIndex = 0
       'txtOut.Text = List1.Text
       'txtCurSong.Text = "1"
'      If ListView1.Text = "" Then
      'List1.RemoveItem List1.ListIndex
'       List1.AddItem "No entries..."
       'txtOut.Text = ""
       'txtCurSong.Text = "0"
       'txtCount.Text = "0"
       'ElseIf List1.Text > "" Then
       'frmMain.Text1.Text = txtOut.Text
       'frmMain.cmdPlaylistT = True
       'Open "c:\Freeplayer.ini" For Output As #1
       'Dim i%
       'For i = 0 To List1.ListCount - 1
       'Print #1, List1.List(i)
 
       'Next
       'Close #1


 '      End If

'List2.Clear
'For n = 0 To List1.ListCount - 1
'durac2 = Len(List1.List(n))
'For i = 0 To durac2
'durac2 = durac2 - 1
'If Mid(List1.List(n), durac2, 1) = "\" Then
'tt.Text = Mid(List1.List(n), durac2 + 1)
'GoTo oper
'End If
'Next
'oper:
            
 '           List2.AddItem tt.Text
  '        List1.ListIndex
          '  List2.ItemData(List2.ListIndex) = List1.ItemData(List1.ListIndex)

'Next
CommonDialog1.FileName = ""
If ListView1.ListItems.count > 13 Then
    ListView1.ColumnHeaders(1).Width = 3171
Else
    ListView1.ColumnHeaders(1).Width = 3411
End If
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ListView1.SetFocus
End Sub

Private Sub Command21_Click()
WindowState = 1
End Sub

Private Sub Command22_Click()
Unload form1
End Sub

Private Sub Command23_Click()
On Error GoTo final
Direction = "left"
duracion = Int(Text10.Text)
Timer6.Interval = duracion * 1
Timer6.Enabled = True
final:
End Sub

Private Sub Command24_Click()
On Error GoTo final
Direction = "right"
duracion = Int(Text10.Text)
Timer6.Interval = duracion * 1
Timer6.Enabled = True
final:
End Sub



















Private Sub Command3_Click()
ListView2.ListItems.Clear
ListView2.ColumnHeaders(1).Width = 3890
playl2.Caption = ""
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ListView2.SetFocus
End Sub

Private Sub Command4_Click()
On Error GoTo endid
Dim FilePath As String
    Dim s As String
    Dim i As Integer
    CommonDialog1.DialogTitle = "Save MGC DJ 2000 Playlist."
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist
    CommonDialog1.Filter = "MGC DJ Playlist (*.djp)|*.djp"

    CommonDialog1.ShowSave

    FilePath = CommonDialog1.FileName
    
    If FilePath = "" Then Exit Sub
        'User pressed cancel
        
    
        For i = 1 To ListView1.ListItems.count
                s = s & ListView1.ListItems(i).Text & vbCrLf & ListView1.ListItems(i).ListSubItems(1).Text & vbCrLf
        Next i
        'Strip of the last vcCrLf
        s = Left(s, Len(s) - 2)
    
        'Everthing is ok creat the list
        Open FilePath For Output As #1
            Print #1, s
        Close #1

    

    'Clear the file name
playl1.Caption = CommonDialog1.FileName
    CommonDialog1.FileName = ""
endid:
End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ListView1.SetFocus
End Sub

Private Sub Command5_Click()
    Dim s As String
     Dim X As Integer
X = 0
ListView1.Sorted = False
    CommonDialog1.DialogTitle = "Load MGC DJ 2000 Playlist."
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    CommonDialog1.Filter = "MGC DJ Playlist (*.djp)|*.djp"

    CommonDialog1.ShowOpen

    FilePath = CommonDialog1.FileName
    
    If FilePath = "" Then Exit Sub
        'User pressed cancel
ListView1.ListItems.Clear
Open FilePath For Input As #1
While Not EOF(1)
Line Input #1, s
If Mid(s, 2, 1) <> ":" Then ListView1.ListItems.Add , , s: X = X + 1: GoTo ira
ListView1.ListItems(X).ListSubItems.Add , , s
ira:
Wend
Close #1
'For i = 1 To lis1.ListItems.Count
'ListView1.ListItems.Add , , lis1.ListItems(i).Text
'ListView1.ListItems(i).ListSubItems.Add , , lis2.ListItems(i)
'Next i
playl1.Caption = CommonDialog1.FileName
CommonDialog1.FileName = ""
If ListView1.ListItems.count > 13 Then
    ListView1.ColumnHeaders(1).Width = 3171
Else
    ListView1.ColumnHeaders(1).Width = 3411
End If
End Sub

Private Sub Command5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ListView1.SetFocus
End Sub

Private Sub Command6_Click()

On Error GoTo endid
Dim FilePath As String
    Dim s As String
    Dim i As Integer
    CommonDialog1.DialogTitle = "Save MGC DJ 2000 Playlist."
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist
    CommonDialog1.Filter = "MGC DJ Playlist (*.djp)|*.djp"

    CommonDialog1.ShowSave

    FilePath = CommonDialog1.FileName
    
    If FilePath = "" Then Exit Sub
        'User pressed cancel
        
    
        For i = 1 To ListView2.ListItems.count
                s = s & ListView2.ListItems(i).Text & vbCrLf & ListView2.ListItems(i).ListSubItems(1).Text & vbCrLf
        Next i
        'Strip of the last vcCrLf
        s = Left(s, Len(s) - 2)
    
        'Everthing is ok creat the list
        Open FilePath For Output As #1
            Print #1, s
        Close #1

    

    'Clear the file name
    playl2.Caption = CommonDialog1.FileName
    CommonDialog1.FileName = ""
endid:
End Sub

Private Sub Command6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ListView2.SetFocus
End Sub

Private Sub Command7_Click()
 Dim s As String
     Dim X As Integer
X = 0
ListView2.Sorted = False
    CommonDialog1.DialogTitle = "Load MGC DJ 2000 Playlist."
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    CommonDialog1.Filter = "MGC DJ Playlist (*.djp)|*.djp"

    CommonDialog1.ShowOpen

    FilePath = CommonDialog1.FileName
    
    If FilePath = "" Then Exit Sub
        'User pressed cancel
ListView2.ListItems.Clear
Open FilePath For Input As #1
While Not EOF(1)
Line Input #1, s
If Mid(s, 2, 1) <> ":" Then ListView2.ListItems.Add , , s: X = X + 1: GoTo ira
ListView2.ListItems(X).ListSubItems.Add , , s
ira:
Wend
Close #1
'For i = 1 To lis1.ListItems.Count
'listview2.ListItems.Add , , lis1.ListItems(i).Text
'listview2.ListItems(i).ListSubItems.Add , , lis2.ListItems(i)
'Next i
playl2.Caption = CommonDialog1.FileName
CommonDialog1.FileName = ""
If ListView2.ListItems.count > 13 Then
    ListView2.ColumnHeaders(1).Width = 3651
Else
    ListView2.ColumnHeaders(1).Width = 3890
End If
End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ListView2.SetFocus
End Sub

Private Sub Command8_Click()
  'working variables
      Dim C As Integer
   Dim q As Integer
   Dim sFile As String
   Dim startStrg As String
   Dim tmp As String
  
  'dim an array to hold the files selected
   Dim FileArray() As String
   ListView2.Sorted = False
   'Timer1.Enabled = False
   'txtTime.Text = ""
   'txtCount.Text = ""
   'Check1.Value = 0
  'set the max buffer large enough to retrieve multiple files
   CommonDialog1.DialogTitle = "Adding song to your playlist." & "  Hold the ctrl key to multi-select."
   CommonDialog1.MaxFileSize = 16384
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "ALL Formats Suported (*.mp3;*.mpg;*.mpeg;*.m1v;*.mp2;*.mpa;*.mp3;*.wav;*.avi;*.asf;*.asx;*.rmi;*.wma;*.wax;*.mid;*.rmi;*.qt;*.aif;*.aifc;*.aiff;*.mov;*.au;*.snd)|*.mp3;*.mpg;*.mpeg;*.m1v;*.mp2;*.mpa;*.mp3;*.wav;*.avi;*.asf;*.asx;*.rmi;*.wma;*.wax;*.mid;*.rmi;*.qt;*.aif;*.aifc;*.aiff;*.mov;*.au;*.snd|MPEG Files (*.mp3;*.mpg;*.mpeg;*.m1v;*.mp2;*.mpa;*.mp3)|*.mp3;*.mpg;*.mpeg;*.m1v;*.mp2;*.mpa;*.mp3|WAV Files (*.wav;*.avi;*.asf;*.asx;*.rmi;*.wma;*.wax)|*.wav;*.avi;*.asf;*.asx;*.rmi;*.wma;*.wax|MIDI Files (*.mid;*.rmi)|*.mid;*.rmi|AIFF Files (*.qt;*.aif;*.aifc;*.aiff;*.mov)|*.qt;*.aif;*.aifc;*.aiff;*.mov|UNIX Files (*.au;*.snd)|*.au;*.snd|ALL Files (*.*)|*.*"
   CommonDialog1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
   CommonDialog1.ShowOpen     ' = 1
  If CommonDialog1.FileName = "" Then Exit Sub
      'assign the string returned from the
  'common dialog to startStrg for further processing.
  'Note that two "nulls" are appended to the
  'end of the string.  This is for use in the StripItem
  'routine below.
   startStrg = CommonDialog1.FileName & Chr(0) & Chr(0)
  
  'Extract each returned filename.
  'If only 1 file was selected, then the string
  'contains the fully-qualified path to the file.
   
  'If more than 1 string file was selected, the
  'string contains the path as the first item,
  'and the FileArray as the rest of the string,
  'each separated by a space.
   For C = 1 To Len(CommonDialog1.FileName)
      
     'extract 1 item from the string
      sFile = StripItem(startStrg, Chr(0))
      
     'if nothing's there, we're done
      If sFile = "" Then Exit For
      
        'redim the filename array
        'to add the new file. FileArray(0) is either the
        'path (if more than 1 file selected), or the
        'fully qualified filename (if only 1 file selected).
         ReDim Preserve FileArray(0 To q)
         FileArray(q) = LCase(sFile)
         
        'increment y by 1 for the next pass
         q = q + 1
      
      Next
      
     'display the results
      Text11.Text = ""
      Text12.Text = ""
      'List1.Clear
      
     'starting with 0, and ending with y-1 (because
     'the above will add 1 more 'y'
     'than there is actual array members)
      For C = 0 To q - 1
      
        'if its the first item, display it in Text1,
        'otherwise display the files selected in Text2.
         If C = 0 Then
               Text11.Text = FileArray(C)
         Else
         tmp = tmp & Text11.Text & "\" & FileArray(C) & vbCrLf
                  If Len(Text11.Text) = 3 Then GoTo falta
         For i = 1 To ListView2.ListItems.count
         If ListView2.ListItems(i).ListSubItems(1).Text = Text11.Text & "\" & FileArray(C) Then GoTo sit
                  Next i

'List1.AddItem Text11.Text & "\" & FileArray(c)
'List2.AddItem FileArray(c)


ListView2.ListItems.Add , , FileArray(C)
ListView2.ListItems(ListView2.ListItems.count).ListSubItems.Add , , Text11.Text & "\" & FileArray(C)

GoTo sit

falta:
For i = 1 To ListView2.ListItems.count
         If ListView2.ListItems(i).ListSubItems(1).Text = Text11.Text & FileArray(C) Then GoTo sit
                  Next i

ListView2.ListItems.Add , , FileArray(C)
ListView2.ListItems(ListView2.ListItems.count).ListSubItems.Add , , Text11.Text & FileArray(C)

'List1.AddItem Text11.Text & FileArray(c)
'List2.AddItem FileArray(c)
sit:
         End If
      Next
      Text12.Text = tmp
       If Text12.Text = "" Then
       For X = 1 To ListView2.ListItems.count
         If ListView2.ListItems(X).ListSubItems(1).Text = Text11.Text Then GoTo tre
                  Next X
       'List1.AddItem Text11.Text
       
        Dim s As String
    Dim Delimiter As Integer
    Dim e As Integer
    
    'Strip off drive letter
    s = Text11.Text

    For e = Len(s) To 0 Step -1
        If Mid(s, e, 1) = "\" Then
            Delimiter = Len(s) - e
            Exit For
        End If
    Next e

    s = Right(s, Delimiter)
        
       
       ListView2.ListItems.Add , , s
ListView2.ListItems(ListView2.ListItems.count).ListSubItems.Add , , Text11.Text

End If
tre:
    
       'If Text10.Text = "" Then Check1.Value = 1
       'txtCount.Text = List1.ListCount
       'If txtCount.Text = "1" Then
       'txtTime.Text = "0"
       'Timer1.Enabled = False
       'ElseIf txtCount.Text > "1" Then
       'txtTime.Text = "0"
       'Timer1.Enabled = True
       'End If
'       List1.ListIndex = 0
       'txtOut.Text = List1.Text
       'txtCurSong.Text = "1"
'      If listview2.Text = "" Then
      'List1.RemoveItem List1.ListIndex
'       List1.AddItem "No entries..."
       'txtOut.Text = ""
       'txtCurSong.Text = "0"
       'txtCount.Text = "0"
       'ElseIf List1.Text > "" Then
       'frmMain.Text1.Text = txtOut.Text
       'frmMain.cmdPlaylistT = True
       'Open "c:\Freeplayer.ini" For Output As #1
       'Dim i%
       'For i = 0 To List1.ListCount - 1
       'Print #1, List1.List(i)
 
       'Next
       'Close #1


 '      End If

'List2.Clear
'For n = 0 To List1.ListCount - 1
'durac2 = Len(List1.List(n))
'For i = 0 To durac2
'durac2 = durac2 - 1
'If Mid(List1.List(n), durac2, 1) = "\" Then
'tt.Text = Mid(List1.List(n), durac2 + 1)
'GoTo oper
'End If
'Next
'oper:
            
 '           List2.AddItem tt.Text
  '        List1.ListIndex
          '  List2.ItemData(List2.ListIndex) = List1.ItemData(List1.ListIndex)

'Next
CommonDialog1.FileName = ""
If ListView2.ListItems.count > 13 Then
    ListView2.ColumnHeaders(1).Width = 3651
Else
    ListView2.ColumnHeaders(1).Width = 3890
End If
End Sub

Private Sub Command8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ListView2.SetFocus
End Sub

Private Sub Command9_Click()
ListView1.ListItems.Clear
ListView1.ColumnHeaders(1).Width = 3411
playl1.Caption = ""
End Sub

Private Sub Command9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ListView1.SetFocus
End Sub

Private Sub contador_Timer()
Dim Pri2 As Single
Dim Pri12 As Single
Dim terc2 As Single
Dim tec12 As Single
Dim seg2 As Integer
Dim seg12 As Integer
Dim Cuart2 As Integer
Dim Cuart12 As Integer
Dim tela2 As String
Dim tela12 As String
Pri2 = MediaPlayer3.Duration / 60
seg2 = Pri2
terc2 = (((Pri2 - seg2) * 100) * 60) / 100
Cuart2 = terc2
If MediaPlayer3.FileName = "" Then Label7.Caption = "0:00 / 0:00": Exit Sub
If Cuart2 < 0 Then seg2 = seg2 - 1: Cuart2 = 60 + Cuart2
Pri12 = MediaPlayer3.CurrentPosition / 60
seg12 = Pri12
terc12 = (((Pri12 - seg12) * 100) * 60) / 100
Cuart12 = terc12
If Cuart12 < 0 Then seg12 = seg12 - 1: Cuart12 = 60 + Cuart12
tela2 = Cuart2
tela12 = Cuart12
If Cuart12 < 10 Then tela12 = "0" & Cuart12
If Cuart2 < 10 Then tela2 = "0" & Cuart2
Label7.Caption = seg12 & ":" & tela12 & " / " & seg2 & ":" & tela2
End Sub

Private Sub cp_Click()
If repeat.Value = 1 Then Timer4.Enabled = False: Timer5.Enabled = False: cp.Caption = "C.P. Off": cp.ForeColor = &HFF&: cpc = &HFF&: repeat.Value = 0: Exit Sub
If repeat.Value = 0 Then Timer4.Enabled = True: Timer5.Enabled = True: cp.Caption = "C.P. On": cp.ForeColor = &HFF00&: cpc = &HFF00&: repeat.Value = 1
End Sub

Private Sub cp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cp.ForeColor = &HFFFFFF
End Sub

Private Sub Dir3_Change()
File3.Path = Dir3.Path
End Sub


Private Sub Drive3_Change()
On Error GoTo checkerror
Dir3.Path = Drive3.Drive
Exit Sub
checkerror:
If (Err.Number = 68) Then
MsgBox "No hay disco en la unidad solicitada !!!."
Drive3.Drive = Dir3.Path
End If
End Sub




Private Sub File3_DblClick()
On Error GoTo solu
dirc = File3.Path
If Len(File3.Path) > 3 Then dirc = File3.Path & "\"
Label8.Caption = File3.FileName
listado2 = File3.ListIndex
MediaPlayer3.FileName = dirc & File3.FileName
Slider7.Max = MediaPlayer3.Duration
Text3.SetFocus
Exit Sub
solu:
Select Case Err.Number
Case 380
If Label14.Caption = "Sp" Then
MsgBox "Archivo imposible de reproducir."
Else
MsgBox "File impossible to reproduce."
End If
Label8.Caption = ""
End Select
End Sub

Private Sub File3_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Key_Enter
On Error GoTo solu
dirc = File3.Path
If Len(File3.Path) > 3 Then dirc = File3.Path & "\"
Label8.Caption = File3.FileName
listado2 = File3.ListIndex
MediaPlayer3.FileName = dirc & File3.FileName
Slider7.Max = MediaPlayer3.Duration
Text3.SetFocus
Case Key_F5
On Error GoTo solu
dirc = File3.Path
If Len(File3.Path) > 3 Then dirc = File3.Path & "\"
Label8.Caption = File3.FileName
listado2 = File3.ListIndex
MediaPlayer3.FileName = dirc & File3.FileName
Slider7.Max = MediaPlayer3.Duration
Text3.SetFocus
End Select
solu:
Select Case Err.Number
Case 380
If Label14.Caption = "Sp" Then
MsgBox "Archivo imposible de reproducir."
Else
MsgBox "File impossible to reproduce."
End If
Label8.Caption = ""
End Select
End Sub

Private Sub Form_DblClick()
If Command21.Visible = False Then Command21.Visible = True: Command22.Visible = True: Label15.Caption = "S": Exit Sub
Command21.Visible = False
Command22.Visible = False
Label15.Caption = "N"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyDivide
On Error GoTo solu
Label4.Caption = ListView1.SelectedItem.Text
MediaPlayer1.FileName = ListView1.SelectedItem.ListSubItems(1).Text
Slider2.Max = MediaPlayer1.Duration
listado = ListView1.SelectedItem.Index
Text1.SetFocus
solu:
Select Case Err.Number
Case 380
If Label14.Caption = "Sp" Then
MsgBox "Formato imposible de reproducir."
Else
MsgBox "File impossible to reproduce."
End If
Label4.Caption = ""
End Select
Case vbKeyMultiply
On Error GoTo solu2
Label5.Caption = ListView2.SelectedItem.Text
MediaPlayer2.FileName = ListView2.SelectedItem.ListSubItems(1).Text
Slider3.Max = MediaPlayer2.Duration
listado1 = ListView2.SelectedItem.Index
Text2.SetFocus
solu2:
Select Case Err.Number
Case 380
If Label14.Caption = "Sp" Then
MsgBox "Formato imposible de reproducir."
Else
MsgBox "File impossible to reproduce."
End If
Exit Sub
End Select
Case vbKeySubtract
On Error GoTo solu3
dirc = File3.Path
If Len(File3.Path) > 3 Then dirc = File3.Path & "\"
Label8.Caption = File3.FileName
listado2 = File3.ListIndex
MediaPlayer3.FileName = dirc & File3.FileName
Slider7.Max = MediaPlayer3.Duration
File3.SetFocus
solu3:
Select Case Err.Number
Case 380
If Label14.Caption = "Sp" Then
MsgBox "Formato imposible de reproducir."
Else
MsgBox "File impossible to reproduce."
End If
Label8.Caption = ""
End Select
Case vbKeyF1:
If one = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = one
Case vbKeyF2:
If two = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = two
Case vbKeyF3:
If three = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = three
Case vbKeyF4:
If four = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = four
Case vbKeyF6:
If five = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = five
Case vbKeyF7:
If six = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = six
Case vbKeyF8:
If seven = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = seven
Case vbKeyF9:
If eight = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = eight
Case vbKeyF10:
If nine = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = nine
Case vbKeyF11:
If ten = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = ten
Case vbKeyF12:
If eleven = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = eleven
End Select
   ShiftKey = Shift And 7
   Select Case ShiftKey
         Case 1
         If KeyCode = vbKeyF1 Then
    If twelve = "                                                                                                    " Then Exit Sub
    MediaPlayer4.FileName = twelve
         End If
         
         If KeyCode = vbKeyF2 Then
    If trece = "                                                                                                    " Then Exit Sub
    MediaPlayer4.FileName = trece
         End If
         
         If KeyCode = vbKeyF3 Then
    If catorce = "                                                                                                    " Then Exit Sub
    MediaPlayer4.FileName = catorce
         End If
         
         If KeyCode = vbKeyF4 Then
    If quince = "                                                                                                    " Then Exit Sub
    MediaPlayer4.FileName = quince
         End If
         
         If KeyCode = vbKeyF6 Then
    If dieciseis = "                                                                                                    " Then Exit Sub
    MediaPlayer4.FileName = dieciseis
         End If
         
         If KeyCode = vbKeyF7 Then
    If diecisiete = "                                                                                                    " Then Exit Sub
    MediaPlayer4.FileName = diecisiete
         End If
         
         If KeyCode = vbKeyF8 Then
    If dieciocho = "                                                                                                    " Then Exit Sub
    MediaPlayer4.FileName = dieciocho
         End If
         
         If KeyCode = vbKeyF9 Then
    If diecinueve = "                                                                                                    " Then Exit Sub
    MediaPlayer4.FileName = diecinueve
         End If
         
         Case 2 ' o vbCtrlMask
On Error GoTo final
Direction = "left"
duracion = Int(Text10.Text)
Timer6.Interval = duracion * 1
Timer6.Enabled = True
final:
         Case 4 ' o vbCtrlMask
On Error GoTo fina
Direction = "right"
duracion = Int(Text10.Text)
Timer6.Interval = duracion * 1
Timer6.Enabled = True
fina:
End Select
End Sub

Private Sub Form_Load()
On Error GoTo fines
Dim comprob As String
comprob = Command$
Dim rgn1 As Long    'main region
    Dim rgn2 As Long    'region to combine with rgn1
    Dim rc As Long        'return code or looping index
ReDim listTemp2(0)
    'irregular form
    rgn1 = CreateEllipticRgn(775, 550, 1, 1)           'create region 1
    'rgn2 = CreateEllipticRgn(400, 0, 450, 50)            'create region 2
    'rc = CombineRgn(rgn1, rgn1, rgn2, RGN_OR)   'add the regions together, the result region is rgn1
    rgn2 = CreateRectRgn(560, 0, 1500, 550)           'create another region
    rc = CombineRgn(rgn1, rgn1, rgn2, RGN_OR) 'subtract rgn2 from rgn1, the result region is rgn1
    rgn2 = CreateEllipticRgn(450, 450, 300, 300)           'create another region
    rc = CombineRgn(rgn1, rgn1, rgn2, RGN_DIFF)   'add rgn2 to rgn1, the result region is rgn1
    rgn2 = CreateEllipticRgn(305, 305, 445, 445)           'create another region
    rc = CombineRgn(rgn1, rgn1, rgn2, RGN_OR)   'add rgn2 to rgn1, the result region is rgn1
        
    rgn2 = CreateRectRgn(0, 0, 220, 550)           'create another region
    rc = CombineRgn(rgn1, rgn1, rgn2, RGN_OR) 'subtract rgn2 from rgn1, the result region is rgn1
        
        'rgn2 = CreateEllipticRgn(69, 29, 79, 34)            'create another region
    'rc = CombineRgn(rgn1, rgn1, rgn2, RGN_DIFF) 'subtract rgn2 from rgn1, the result region is rgn1
    rc = SetWindowRgn(Me.hWnd, rgn1, True)

Move (Screen.Width - form1.Width) / 2, (Screen.Height - form1.Height) / 2
'Call ExplodeForm(Me, 100)
KeyPreview = True
ListView1.ColumnHeaders(1).Width = 3411
ListView2.ColumnHeaders(1).Width = 3890
listado = 1
listado1 = 1
Slider1.Value = -3000
MediaPlayer1.Volume = 0
MediaPlayer4.Volume = 0
MediaPlayer2.Volume = -6000
Timer7.Enabled = False
req = 0
bec = &HFF&
mec = &HFF00&
smc = &HFF00&
cpc = &HFF00&
Slider2.Max = 1
Slider3.Max = 1
Slider7.Max = 1
Slider4.Value = 0
Slider5.Value = 0
Slider6.Value = 0
Vol1.Value = -3000
Vol2.Value = 0
Vol1.Enabled = False
Vol2.Enabled = False
Slider8.Value = -6000
Slider10.Value = -6000
Slider11.Value = -3000
Temp = 0
durac = 0

  WaveReset
    
    Rate = CLng(GetSetting("MGC DJ 2000", "StartUp", "Rate", "110025"))
    Channels = CInt(GetSetting("MGC DJ 2000", "StartUp", "Channels", "1"))
    Resolution = CInt(GetSetting("MGC DJ 2000", "StartUp", "Resolution", "16"))
    WaveFileName = GetSetting("MGC DJ 2000", "StartUp", "WaveFileName", "C:\Radio.wav")
    WaveAutomaticSave = GetSetting("MGC DJ 2000", "StartUp", "WaveAutomaticSave", "True")

    WaveRecordingImmediate = True
    WaveRecordingReady = False
    WaveRecording = False
    WavePlaying = False
    
    'Be sure to change the Value property of the appropriate button!!
    'if you change the default values!
    
    WaveSet
    'WaveRecordingStartTime = Now + TimeSerial(0, 15, 0)
    'WaveRecordingStopTime = WaveRecordingStartTime + TimeSerial(0, 15, 0)
    WaveMidiFileName = ""
    WaveRenameNecessary = False
    
Dim Leer As MGClectura
Dim Registros As Integer
Dim Lon As Integer
numeroDeCanal = FreeFile
Open App.Path & "/mgcdj2000.cfg" For Random As #numeroDeCanal Len = 434
Registros = LOF(numeroDeCanal) / 434
For i = 1 To Registros
Get #numeroDeCanal, i, Leer
playl1.Caption = RTrim(Leer.playl11)
playl2.Caption = RTrim(Leer.playl22)
Drive3.Drive = Leer.Unidad3
Dir3.Path = Leer.Ruta3
Label6.Caption = RTrim(Leer.BackSelect)
Label10.Caption = Leer.Inisound
Label14.Caption = Leer.Languaje
Label15.Caption = Leer.Mini
Label16.Caption = RTrim(Leer.djeffect)
Labeldj.Caption = Leer.dtext
Next i
solut:
Close #numeroDeCanal
If Label10.Caption = "S" Then
    MediaPlayer4.FileName = App.Path & "\mgcdj2000sound.mp3"
ElseIf Label10.Caption = "N" Then GoTo xtt
    Else
    MediaPlayer4.FileName = App.Path & "\mgcdj2000sound.mp3"
End If
xtt:
If Label15.Caption = "N" Then Command21.Visible = False: Command22.Visible = False
Dim dssa As String
dssa = Trim(Labeldj.Caption)
If dssa = "" Then tete = "Martín Caleau": GoTo neos
tete = Labeldj.Caption
neos:

If Label6.Caption <> "" Then
On Error GoTo specialskinerr

Dim skiner As Skin
Dim Regskin As Integer
numCanal = FreeFile
Open App.Path & "\skins\" & Label6.Caption For Random As #numCanal Len = 264
Regskin = LOF(numCanal) / 264
For i = 1 To Regskin
Get #numCanal, i, skiner

colorf1 = RTrim(skiner.form1)
     colorf1exmin = RTrim(skiner.ME_B)
     colorf1effect = RTrim(skiner.Effect_B)
     colorf1play = RTrim(skiner.Play_B)
     colorf1stop = RTrim(skiner.Stop_B)
     colorf1pause = RTrim(skiner.Pause_B)

colorf2 = RTrim(skiner.form2)
colorf3 = RTrim(skiner.form3)
colorf4 = RTrim(skiner.Form4)
colorf6 = RTrim(skiner.Form6)
colormixer = RTrim(skiner.mixer)
colorfsettings = RTrim(skiner.fsetting)
coloremail = RTrim(skiner.email)
colordirname = RTrim(skiner.dirname)

Next i
Close #numCanal

form1.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1)
form1.Command21.BackColor = colorf1exmin
form1.Command22.BackColor = colorf1exmin
form1.Label36.ForeColor = colorf1exmin
form1.Image4.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1play)
form1.Image5.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1stop)
form1.Image6.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1pause)
form1.Image3.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1play)
form1.Image2.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1stop)
form1.Image1.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1pause)
form1.B_Play.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1play)
form1.B_Stop.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1stop)
form1.B_Pause.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1pause)
form1.Image7.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image8.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image9.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image10.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image11.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image12.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image13.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image14.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image15.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image16.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image17.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image18.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image19.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image20.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image21.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image22.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image23.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image24.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image25.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image29.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image27.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image26.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image30.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image31.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image28.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)
form1.Image32.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf1effect)

End If

specialskinerr:

If Label14.Caption = "Sp" Then
Drive3.ToolTipText = "Unidad Music Special"
File3.ToolTipText = "Music Box Special / Haga Doble Click o presione F5 o ENTER"
ListView1.ToolTipText = "Music Box A / Haga Doble Click o presione F5 o ENTER"
ListView2.ToolTipText = "Music Box B / Haga Doble Click o presione F5 o ENTER"
Dir3.ToolTipText = "Directorio Music Special"
Text1.ToolTipText = "Buscador Music Box A / Presiona la tecla ´F5´ para reproducir la canción seleccionada"
Text2.ToolTipText = "Buscador Music Box B / Presiona la tecla ´F5´ para reproducir la canción seleccionada"
Text3.ToolTipText = "Buscador Music Box Special / Presiona la tecla ´F5´ para reproducir la canción seleccionada"
Slider2.ToolTipText = "Posición de Music A"
Slider3.ToolTipText = "Posición de Music B"
Slider7.ToolTipText = "Posición de Music Special"
Vol1.ToolTipText = "Volumen A"
Vol2.ToolTipText = "Volumen B"
Slider8.ToolTipText = "Volumen Special"
Slider10.ToolTipText = "Volumen Efecto DJ"
Image32.ToolTipText = "Parar Efecto DJ"
B_Play.ToolTipText = "Reproducir A"
B_Stop.ToolTipText = "Parar A"
B_Pause.ToolTipText = "Pausar A"
Image3.ToolTipText = "Reproducir B"
Image2.ToolTipText = "Parar B"
Image1.ToolTipText = "Pausar B"
Image4.ToolTipText = "Reproducir Special"
Image5.ToolTipText = "Parar Special"
Image6.ToolTipText = "Pausar Special"
Label2.ToolTipText = "Duración Music A"
Label3.ToolTipText = "Duración Music B"
Label7.ToolTipText = "Duración Special"
Label9.Caption = "Efectos DJ"
Text4.ToolTipText = "Efecto Minutos A / Seleccione el tiempo deseado en minutos en que comience la canción y luego presiona ´Play´"
Text5.ToolTipText = "Efecto Segundos A  / Seleccione el tiempo deseado en segundos en que comience la canción y luego presiona ´Play´"
Text7.ToolTipText = "Efecto Minutos B / Seleccione el tiempo deseado en minutos en que comience la canción y luego presiona ´Play´"
Text6.ToolTipText = "Efecto Segundos B  / Seleccione el tiempo deseado en segundos en que comience la canción y luego presiona ´Play´"
Text9.ToolTipText = "Efecto Minutos Special / Seleccione el tiempo deseado en minutos en que comience la canción y luego presiona ´Play´"
Text8.ToolTipText = "Efecto Segundos Special  / Seleccione el tiempo deseado en segundos en que comience la canción y luego presiona ´Play´"
Label1.ToolTipText = "Hora Actual"
Label11.ToolTipText = "Doble Click: Borra Efecto Min / Seg A"
Label12.ToolTipText = "Doble Click: Borra Efecto Min / Seg B"
Label13.ToolTipText = "Doble Click: Borra Efecto Min / Seg Special"
Label39.Caption = "Efecto  Mixer On"
Label38.Caption = "Efecto Balance Off"
Label39.ToolTipText = "Permite mixar canciones entre el Playlist A y B"
Label17.ToolTipText = "Presiona F1"
Label18.ToolTipText = "Presiona F2"
Label19.ToolTipText = "Presiona F3"
Label20.ToolTipText = "Presiona F4"
Label21.ToolTipText = "Presiona F6"
Label22.ToolTipText = "Presiona F7"
Label23.ToolTipText = "Presiona F8"
Label24.ToolTipText = "Presiona F9"
Label25.ToolTipText = "Presiona F10"
Label26.ToolTipText = "Presiona F11"
Label27.ToolTipText = "Presiona F12"
Label28.ToolTipText = "Presiona Shift + F1"
Label29.ToolTipText = "Presiona Shift + F2"
Label30.ToolTipText = "Presiona Shift + F3"
Label31.ToolTipText = "Presiona Shift + F4"
Label32.ToolTipText = "Presiona Shift + F6"
Label33.ToolTipText = "Presiona Shift + F7"
Label34.ToolTipText = "Presiona Shift + F8"
Label35.ToolTipText = "Presiona Shift + F9"
Slider1.ToolTipText = "Efecto Mixer Scroll (Funciona solo si Efecto Mixer está en ON)"
Command21.ToolTipText = "Minimizar"
Command22.ToolTipText = "Salir"
Command23.ToolTipText = "Auto Mix Efecto Izquierda / Presiona la tecla ´Control´ izquierda"
Command24.ToolTipText = "Auto Mix Efecto Derecha / Presiona la tecla ´Alt´ izquierda"
Text10.ToolTipText = "Seleccione la Velocidad del Mix (1-99)"
form1.Label41.Caption = "Grabar"
form1.Label41.ToolTipText = "Graba en Tiempo Real todo lo reproducido en su equipo (Incluyendo Mix, dj effects, mic, midis, etc)."
form1.Label44.Caption = "Guardar"
form1.Label44.ToolTipText = "Guarda en un archivo .wav lo que Ud. grabó."
form1.Label45.Caption = "Configurar"
form1.Label45.ToolTipText = "Configure Calidad de Grabación, modo, tempo, midis karaoke, etc."
form1.Label43.ToolTipText = "Reproduce lo grabado."
form1.Label40.ToolTipText = "Borra todo lo grabado. Reestablece Configuración Inicial."
form1.Label42.ToolTipText = "Para la grabación o lo que se está reproduciendo."
Label36.Caption = "OPCIONES"
Label36.ToolTipText = "Menu de Opciones (Haga Click en el Botón Derecho del Mouse)."
form1.StatisticsLabel.ToolTipText = "Información sobre la grabación."
form1.Sliderxx.ToolTipText = "Seleccione el inicio para reproducir."
'Slider11.ToolTipText = "Volumen Master"
MGC.ToolTipText = "Sección Texto Dj"
Command2.ToolTipText = "Agregar Archivo Musical al Playlist A"
Command8.ToolTipText = "Agregar Archivo Musical al Playlist B"
Command5.ToolTipText = "Cargar MGC DJ Playlist A"
Command7.ToolTipText = "Cargar MGC DJ Playlist B"
Command4.ToolTipText = "Guardar MGC DJ Playlist A"
Command6.ToolTipText = "Guardar MGC DJ Playlist B"
Command9.ToolTipText = "Crear Nuevo MGC DJ Playlist A"
Command3.ToolTipText = "Crear Nuevo MGC DJ Playlist B"
Label37.ToolTipText = "Selecciona el Modo de Mix. Funciona si el Efecto Mixer está en ON, utilizando el Efecto Mixer Scroll"
cp.ToolTipText = "Continuar Reproduciendo la canción siguiente en los playlists al finalizar la actual"
'Check1.Caption = "Efecto  Mixer On"
End If
Label36.ForeColor = Command21.BackColor


Dim XXX As Sonidos
Dim Registros2 As Integer
Dim ok2 As String
numeroDeCan5 = FreeFile
If Label16.Caption = "" Then Label16.Caption = App.Path & "/mgcustomize.dje"
ok2 = RTrim(form1.Label16.Caption)
Open ok2 For Random As #numeroDeCan5 Len = 2128
Registros2 = LOF(numeroDeCan5) / 2128
For i = 1 To Registros2
Get #numeroDeCan5, i, XXX
one = XXX.one
two = XXX.two
three = XXX.three
four = XXX.four
five = XXX.five
six = XXX.six
seven = XXX.seven
eight = XXX.eight
nine = XXX.nine
ten = XXX.ten
eleven = XXX.eleven
twelve = XXX.twelve
trece = XXX.trece
catorce = XXX.catorce
quince = XXX.quince
dieciseis = XXX.dieciseis
diecisiete = XXX.diecisiete
dieciocho = XXX.dieciocho
diecinueve = XXX.diecinueve
Label17.Caption = RTrim(XXX.c1)
Label18.Caption = RTrim(XXX.c2)
Label19.Caption = RTrim(XXX.c3)
Label20.Caption = RTrim(XXX.c4)
Label21.Caption = RTrim(XXX.c5)
Label22.Caption = RTrim(XXX.c6)
Label23.Caption = RTrim(XXX.c7)
Label24.Caption = RTrim(XXX.c8)
Label25.Caption = RTrim(XXX.c9)
Label26.Caption = RTrim(XXX.c10)
Label27.Caption = RTrim(XXX.c11)
Label28.Caption = RTrim(XXX.c12)
Label29.Caption = RTrim(XXX.c13)
Label30.Caption = RTrim(XXX.c14)
Label31.Caption = RTrim(XXX.c15)
Label32.Caption = RTrim(XXX.c16)
Label33.Caption = RTrim(XXX.c17)
Label34.Caption = RTrim(XXX.c18)
Label35.Caption = RTrim(XXX.c19)
Next i
Close #numeroDeCan5
'If IsConnected = True Then
'tupdate.Enabled = True
'    End If
If comprob = "" Then GoTo yeta
MediaPlayer4.Stop
    Dim e As Integer
    Dim Delimiter As String
      
            For e = Len(comprob) To 0 Step -1
        If Mid(comprob, e, 1) = "\" Then
            Delimiter = Len(comprob) - e
            Exit For
        End If
    Next e

    Delimiter = Right(comprob, Delimiter)

    ListView1.ListItems.Add , , Delimiter
    ListView1.ListItems(ListView1.ListItems.count).ListSubItems.Add , , comprob

Label4.Caption = ListView1.SelectedItem.Text
'listado = File1.ListIndex
MediaPlayer1.FileName = ListView1.ListItems(ListView1.ListItems.count).ListSubItems(1).Text
Slider2.Max = MediaPlayer1.Duration
'Text1.SetFocus
If Label4.Caption = ListView1.SelectedItem.Text Then GoTo noplay1
yeta:
Dim s As String
Dim X As Integer
ListView1.ListItems.Clear
If playl1.Caption = "" Then GoTo noplay1
Open playl1.Caption For Input As #1
While Not EOF(1)
Line Input #1, s
If Mid(s, 2, 1) <> ":" Then ListView1.ListItems.Add , , s: X = X + 1: GoTo ira
ListView1.ListItems(X).ListSubItems.Add , , s
ira:
Wend
Close #1
If ListView1.ListItems.count > 13 Then
    ListView1.ColumnHeaders(1).Width = 3171
Else
    ListView1.ColumnHeaders(1).Width = 3411
End If
noplay1:
Dim s2 As String
Dim X2 As Integer
On Error GoTo fines
ListView2.ListItems.Clear
If playl2.Caption = "" Then GoTo fines
Open playl2.Caption For Input As #2
While Not EOF(2)
Line Input #2, s2
If Mid(s2, 2, 1) <> ":" Then ListView2.ListItems.Add , , s2: X2 = X2 + 1: GoTo ira2
ListView2.ListItems(X2).ListSubItems.Add , , s2
ira2:
Wend
Close #1
If ListView2.ListItems.count > 13 Then
    ListView2.ColumnHeaders(1).Width = 3651
Else
    ListView2.ColumnHeaders(1).Width = 3890
End If
Exit Sub
fines:
'MsgBox Err.Number
On Error GoTo final
'Kill (Form1.Label16.Caption)
Dim XXX9 As Sonidos
Dim Registros9 As Integer
numeroDeCan23 = FreeFile
XXX9.one = App.Path & "\Samples\mgc-presents.mp3"
XXX9.two = App.Path & "\Samples\mgc-goodnight.mp3"
XXX9.three = App.Path & "\Samples\mgc-radionew.mp3"
XXX9.four = App.Path & "\Samples\mgc-lallama.mp3"
XXX9.five = App.Path & "\Samples\mgc-smile.mp3"
XXX9.c1 = "mgcdj-1"
XXX9.c2 = "mgcdj-2"
XXX9.c3 = "mgcdj-3"
XXX9.c4 = "mgcdj-4"
XXX9.c5 = "mgcdj-5"
form1.Label16.Caption = App.Path & "/mgcustomize.dje"
Open form1.Label16.Caption For Random As #numeroDeCan23 Len = 2128
Registros9 = LOF(numeroDeCan23) / 2128
Registros9 = Registros9 + 1
Put #numeroDeCan23, Registros9, XXX9
Close #numeroDeCan23
one = App.Path & "\Samples\mgc-presents.mp3"
two = App.Path & "\Samples\mgc-goodnight.mp3"
three = App.Path & "\Samples\mgc-radionew.mp3"
four = App.Path & "\Samples\mgc-lallama.mp3"
five = App.Path & "\Samples\mgc-smile.mp3"
form1.Label17.Caption = "mgcdj-1"
form1.Label18.Caption = "mgcdj-2"
form1.Label19.Caption = "mgcdj-3"
form1.Label20.Caption = "mgcdj-4"
form1.Label21.Caption = "mgcdj-5"
final:
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mousedown = Button
    If Button = 1 Then
       Dim ReturnVal As Long
       X = ReleaseCapture()
       ReturnVal = SendMessage(hWnd, WM_NCLBUTTONDOWN, 2, 0)
   
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label36.ForeColor = Command21.BackColor
Label38.ForeColor = bec
Label39.ForeColor = mec
Label37.ForeColor = smc
Label40.ForeColor = &H0&
Label41.ForeColor = &H0&
Label42.ForeColor = &H0&
Label43.ForeColor = &H0&
Label44.ForeColor = &H0&
Label45.ForeColor = &H0&
cp.ForeColor = cpc
 If Check4.Value = 1 Then
 If Movetest = 0 Then
                mgcmixer.Move form1.Left + 3350, form1.Top + 400
        End If
        End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mousedown = NO_BUTTON
End Sub

Private Sub Form_Unload(Cancel As Integer)
MediaPlayer1.Stop
MediaPlayer2.Stop
MediaPlayer3.Stop
MediaPlayer4.Stop
    
'Call ImplodeForm(Me, 1, 200, 2)

On Error GoTo other
Kill (App.Path & "/mgcdj2000.cfg")
other:
Dim Leer As MGClectura
Dim Registros As Integer
numeroDeCana = FreeFile
Leer.playl11 = playl1.Caption
Leer.playl22 = playl2.Caption
Leer.Unidad3 = Drive3.Drive
Leer.Ruta3 = Dir3.Path
Leer.BackSelect = Label6.Caption
Leer.Inisound = Label10.Caption
Leer.Languaje = Label14.Caption
Leer.Mini = Label15.Caption
Leer.djeffect = Label16.Caption
Leer.dtext = Labeldj.Caption
Open App.Path & "/mgcdj2000.cfg" For Random As #numeroDeCana Len = 434
Registros = LOF(numeroDeCana) / 434
Registros = Registros + 1
Put #numeroDeCana, Registros, Leer
Close #numeroDeCana
WaveClose
    Call SaveSetting("MGC DJ 2000", "StartUp", "Rate", CStr(Rate))
    Call SaveSetting("MGC DJ 2000", "StartUp", "Channels", CStr(Channels))
    Call SaveSetting("MGC DJ 2000", "StartUp", "Resolution", CStr(Resolution))
    Call SaveSetting("MGC DJ 2000", "StartUp", "WaveFileName", WaveFileName)
    Call SaveSetting("MGC DJ 2000", "StartUp", "WaveAutomaticSave", CStr(WaveAutomaticSave))
    If WaveRenameNecessary Then
        Name WaveShortFileName As WaveLongFileName
        WaveRenameNecessary = False
        WaveShortFileName = ""
    End If
If req = 1 Then
Exit Sub
Else: End
End If
End Sub

Private Sub Image1_Click()
On Error GoTo bye
MediaPlayer2.Pause
bye:
End Sub

Private Sub Image2_Click()
MediaPlayer2.Stop
MediaPlayer2.CurrentPosition = 0
End Sub

Private Sub Image3_Click()
On Error GoTo bye
MediaPlayer2.Play
If Text7.Text = "0" And Text6.Text = "0" Then Exit Sub
MediaPlayer2.CurrentPosition = Text7.Text * 60 + Text6.Text
bye:
End Sub

Private Sub Image32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MediaPlayer4.Stop
MediaPlayer4.CurrentPosition = 0
Image32.BorderStyle = 1
End Sub

Private Sub Image32_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image32.BorderStyle = 0
End Sub

Private Sub Image4_Click()
On Error GoTo bye
MediaPlayer3.Play
If Text9.Text = "0" And Text8.Text = "0" Then Exit Sub
MediaPlayer3.CurrentPosition = Text9.Text * 60 + Text8.Text
bye:
End Sub

Private Sub Image5_Click()
MediaPlayer3.Stop
MediaPlayer3.CurrentPosition = 0
End Sub

Private Sub Image6_Click()
On Error GoTo bye
MediaPlayer3.Pause
bye:
End Sub





Private Sub Label11_DblClick()
Text5.Text = "0"
Text4.Text = "0"
End Sub

Private Sub Label12_DblClick()
Text6.Text = "0"
Text7.Text = "0"
End Sub

Private Sub Label13_DblClick()
Text8.Text = "0"
Text9.Text = "0"
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.BorderStyle = 1
If one = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = one
End Sub

Private Sub Label17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.BorderStyle = 0
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.BorderStyle = 1
If two = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = two
End Sub

Private Sub Label18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.BorderStyle = 0
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.BorderStyle = 1
If three = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = three
End Sub

Private Sub Label19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.BorderStyle = 0
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.BorderStyle = 1
If four = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = four
End Sub

Private Sub Label20_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.BorderStyle = 0
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.BorderStyle = 1
If five = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = five
End Sub

Private Sub Label21_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.BorderStyle = 0
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image12.BorderStyle = 1
If six = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = six
End Sub

Private Sub Label22_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image12.BorderStyle = 0
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image13.BorderStyle = 1
If seven = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = seven
End Sub

Private Sub Label23_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image13.BorderStyle = 0
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image14.BorderStyle = 1
If eight = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = eight
End Sub

Private Sub Label24_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image14.BorderStyle = 0
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image15.BorderStyle = 1
If nine = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = nine
End Sub

Private Sub Label25_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image15.BorderStyle = 0
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image16.BorderStyle = 1
If ten = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = ten
End Sub

Private Sub Label26_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image16.BorderStyle = 0
End Sub

Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image17.BorderStyle = 1
If eleven = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = eleven
End Sub

Private Sub Label27_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image17.BorderStyle = 0
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image18.BorderStyle = 1
If twelve = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = twelve
End Sub

Private Sub Label28_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image18.BorderStyle = 0
End Sub

Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image19.BorderStyle = 1
If trece = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = trece
End Sub

Private Sub Label29_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image19.BorderStyle = 0
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image20.BorderStyle = 1
If catorce = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = catorce
End Sub

Private Sub Label30_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image20.BorderStyle = 0
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image21.BorderStyle = 1
If quince = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = quince
End Sub

Private Sub Label31_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image21.BorderStyle = 0
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.BorderStyle = 1
If dieciseis = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = dieciseis
End Sub

Private Sub Label32_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.BorderStyle = 0
End Sub

Private Sub Label33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image23.BorderStyle = 1
If diecisiete = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = diecisiete
End Sub

Private Sub Label33_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image23.BorderStyle = 0
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image24.BorderStyle = 1
If dieciocho = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = dieciocho
End Sub

Private Sub Label34_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image24.BorderStyle = 0
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image25.BorderStyle = 1
If diecinueve = "                                                                                                    " Then Exit Sub
MediaPlayer4.FileName = diecinueve
End Sub

Private Sub Label35_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image25.BorderStyle = 0
End Sub

Private Sub Label36_Click()
form2.Show
End Sub

Private Sub Label36_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label36.ForeColor = &HFFFFFF
End Sub

Private Sub Label36_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label36.ForeColor = &HFFFFFF
End Sub

Private Sub Label36_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label36.ForeColor = Command21.BackColor
End Sub

Private Sub Label37_Click()
If Check3.Value = 0 Then Label37.Caption = "Mix Mode 1": Label37.ForeColor = &HFF00&: smc = &HFF00&: Check3.Value = 1: Exit Sub
Label37.Caption = "Mix Mode 2"
Label37.ForeColor = &HFF&
smc = &HFF&
Check3.Value = 0
End Sub

Private Sub Label37_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label37.ForeColor = &HFFFFFF
End Sub

Private Sub Label38_Click()
If Check2.Value = 0 Then
Label38.ForeColor = &HFF00&
Check2.Value = 1
bec = &HFF00&
If Label14.Caption = "Sp" Then
Label38.Caption = "Efecto Balance On"
Else
Label38.Caption = "Balance Effect On"
End If
Slider6.SetFocus
Else
Label38.ForeColor = &HFF&
Check2.Value = 0
bec = &HFF&
If Label14.Caption = "Sp" Then
Label38.Caption = "Efecto Balance Off"
Else
Label38.Caption = "Balance Effect Off"
End If
Slider4.Value = MediaPlayer1.Balance
Slider5.Value = MediaPlayer2.Balance
End If
End Sub

Private Sub Label38_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label38.ForeColor = &HFFFFFF
End Sub

Private Sub Label39_Click()
If Check1.Value = 0 Then
Label39.ForeColor = &HFF00&
Check1.Value = 1
mec = &HFF00&
If Label14.Caption = "Sp" Then
Label39.Caption = "Efecto  Mixer On"
Else
Label39.Caption = "Mixer  Effect On"
End If
Vol1.Enabled = False
Vol2.Enabled = False
Slider1.SetFocus
Else
Label39.ForeColor = &HFF&
Check1.Value = 0
mec = &HFF&
If Label14.Caption = "Sp" Then
Label39.Caption = "Efecto  Mixer Off"
Else
Label39.Caption = "Mixer  Effect Off"
End If
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Vol1.Enabled = True
Vol2.Enabled = True
End If
End Sub

Private Sub Label39_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label39.ForeColor = &HFFFFFF
End Sub

Private Sub Label40_Click()
  Sliderxx.Max = 10
    Sliderxx.Value = 0
    Sliderxx.Refresh
    Label41.Enabled = True
    Label42.Enabled = False
    Label43.Enabled = False
    Label44.Enabled = False
    
    WaveReset
    
    Rate = CLng(GetSetting("MGC DJ 2000", "StartUp", "Rate", "110025"))
    Channels = CInt(GetSetting("MGC DJ 2000", "StartUp", "Channels", "1"))
    Resolution = CInt(GetSetting("MGC DJ 2000", "StartUp", "Resolution", "16"))
    WaveFileName = GetSetting("MGC DJ 2000", "StartUp", "WaveFileName", "C:\Radio.wav")
    WaveAutomaticSave = GetSetting("MGC DJ 2000", "StartUp", "WaveAutomaticSave", "True")

    WaveRecordingImmediate = True
    WaveRecordingReady = False
    WaveRecording = False
    WavePlaying = False
    WaveMidiFileName = ""
    'Be sure to change the Value property of the appropriate button!!
    'if you change the default values!
    
    WaveSet
    If WaveRenameNecessary Then
        Name WaveShortFileName As WaveLongFileName
        WaveRenameNecessary = False
        WaveShortFileName = ""
    End If
End Sub

Private Sub Label40_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label40.ForeColor = &HFFFFFF
Label41.ForeColor = &H0&
Label42.ForeColor = &H0&
Label43.ForeColor = &H0&
Label44.ForeColor = &H0&
Label45.ForeColor = &H0&
End Sub

Private Sub Label41_Click()
  Dim settings As String
    Dim Alignment As Integer
      
    Alignment = Channels * Resolution / 8
    
    settings = "set capture alignment " & CStr(Alignment) & " bitspersample " & CStr(Resolution) & " samplespersec " & CStr(Rate) & " channels " & CStr(Channels) & " bytespersec " & CStr(Alignment * Rate)
    WaveReset
    WaveSet
    WaveRecord
    WaveRecordingStartTime = Now
    Label42.Enabled = True   'Enable the STOP BUTTON
    Label43.Enabled = False  'Disable the "PLAY" button
    Label44.Enabled = False  'Disable the "SAVE AS" button
    Label41.Enabled = False 'Disable the "RECORD" button
End Sub

Private Sub Label41_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label40.ForeColor = &H0&
Label41.ForeColor = &HFFFFFF
Label42.ForeColor = &H0&
Label43.ForeColor = &H0&
Label44.ForeColor = &H0&
Label45.ForeColor = &H0&
End Sub

Private Sub Label42_Click()
 WaveStop
    Label44.Enabled = True  'Enable the "SAVE AS" button
    Label43.Enabled = True  'Enable the "PLAY" button
    Label42.Enabled = False 'Disable the "STOP" button
    If WavePosition = 0 Then
        Sliderxx.Max = 10
    Else
        If WaveRecordingImmediate And (Not WavePlaying) Then Sliderxx.Max = WavePosition
        If (Not WaveRecordingImmediate) And WaveRecording Then Sliderxx.Max = WavePosition
    End If
    If WaveRecording Then WaveRecordingReady = True
    WaveRecordingStopTime = Now
    WaveRecording = False
    WavePlaying = False
    frmSettings.optRecordProgrammed.Value = False
    frmSettings.optRecordImmediate.Value = True
    frmSettings.lblTimes.Visible = False
End Sub

Private Sub Label42_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label40.ForeColor = &H0&
Label41.ForeColor = &H0&
Label42.ForeColor = &HFFFFFF
Label43.ForeColor = &H0&
Label44.ForeColor = &H0&
Label45.ForeColor = &H0&
End Sub

Private Sub Label43_Click()
  WavePlayFrom (Sliderxx.Value)
    WavePlaying = True
    Label42.Enabled = True
    Label43.Enabled = False
End Sub

Private Sub Label43_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label40.ForeColor = &H0&
Label41.ForeColor = &H0&
Label42.ForeColor = &H0&
Label43.ForeColor = &HFFFFFF
Label44.ForeColor = &H0&
Label45.ForeColor = &H0&
End Sub

Private Sub Label44_Click()
   Dim sName As String
    
    If WaveMidiFileName = "" Then
        sName = "Radio_from_" & CStr(WaveRecordingStartTime) & "_to_" & CStr(WaveRecordingStopTime)
        sName = Replace(sName, ":", "-")
        sName = Replace(sName, " ", "_")
        sName = Replace(sName, "/", "-")
    Else
        sName = WaveMidiFileName
        sName = Replace(sName, "MID", "wav")
    End If
  
    CommonDialogxx.FileName = sName
    CommonDialogxx.CancelError = True
    On Error GoTo ErrHandler1
    CommonDialogxx.Filter = "WAV file (*.wav*)|*.wav"
    CommonDialogxx.Flags = &H2 Or &H400
    CommonDialogxx.ShowSave
    sName = CommonDialogxx.FileName
    
    WaveSaveAs (sName)
    Exit Sub
ErrHandler1:
End Sub

Private Sub Label44_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label40.ForeColor = &H0&
Label41.ForeColor = &H0&
Label42.ForeColor = &H0&
Label43.ForeColor = &H0&
Label44.ForeColor = &HFFFFFF
Label45.ForeColor = &H0&
End Sub

Private Sub Label45_Click()
Dim strWhat As String
    ' show the user entry form modally
    If form1.Label14.Caption = "Sp" Then strWhat = MsgBox("Si continua, perderá todos los datos guardados en memoria!", vbOKCancel): GoTo english
    strWhat = MsgBox("If you continue your data will be lost!", vbOKCancel)
english:
    If strWhat = vbCancel Then
        Exit Sub
        
        End If
    Sliderxx.Max = 10
    Sliderxx.Value = 0
    Sliderxx.Refresh
    Label41.Enabled = True
    Label42.Enabled = False
    Label43.Enabled = False
    Label44.Enabled = False
    
    WaveReset
    
    Rate = CLng(GetSetting("MGC DJ 2000", "StartUp", "Rate", "110025"))
    Channels = CInt(GetSetting("MGC DJ 2000", "StartUp", "Channels", "1"))
    Resolution = CInt(GetSetting("MGC DJ 2000", "StartUp", "Resolution", "16"))
    WaveFileName = GetSetting("MGC DJ 2000", "StartUp", "WaveFileName", "C:\Radio.wav")
    WaveAutomaticSave = GetSetting("MGC DJ 2000", "StartUp", "WaveAutomaticSave", "True")

    WaveRecordingImmediate = True
    WaveRecordingReady = False
    WaveRecording = False
    WavePlaying = False
    
    'Be sure to change the Value property of the appropriate button!!
    'if you change the default values!
    
    WaveSet
    frmSettings.optRecordImmediate.Value = True
    Unload frmSettings
    Load frmSettings
    frmSettings.Show 'vbModal
End Sub

Private Sub Label45_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label40.ForeColor = &H0&
Label41.ForeColor = &H0&
Label42.ForeColor = &H0&
Label43.ForeColor = &H0&
Label44.ForeColor = &H0&
Label45.ForeColor = &HFFFFFF
End Sub

Private Sub listView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   ListView1.MultiSelect = False
'   Set nodX = ListView1.SelectedItem ' Elemento arrastrado.
End Sub

Private Sub listView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'      If Button = vbLeftButton Then ' Indica una operación de arrastrar.
'           indrag = True ' Establece el indicador como verdadero.
' ListView1.Drag vbBeginDrag
       ' Establece el icono de arrastre con el método CreateDragImage.
'      ListView1.DragIcon = ListView1.SelectedItem.CreateDragImage
      ' Operación de arrastre.
'     End If
End Sub

Private Sub ListView1_DragDrop(Source As Control, X As Single, Y As Single)
 '  If ListView1.DropHighlight Is Nothing Then
 '     Set ListView1.DropHighlight = Nothing
 '     indrag = False
 '     ListView1.MultiSelect = True
 '     Exit Sub
 '  Else
 '     If nodX = ListView1.DropHighlight Then ListView1.MultiSelect = True: Exit Sub
 '     Cls
 '     ListView1.Drag vbEndDrag
 '     MsgBox nodX.Text & " colocado en " & ListView1.DropHighlight.Text
 '     Set ListView1.DropHighlight = Nothing
  '    indrag = False
  ' ListView1.MultiSelect = True
 '  End If
End Sub

Private Sub listView1_DragOver(Source As Control, X As Single, Y As Single, state As Integer)
'      If indrag = True Then
      
       ' Establece las coordenadas del mouse en DropHighlight.
 '     Set ListView1.DropHighlight = ListView1.HitTest(X, Y)
'   End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next

    ListView1.SortKey = ColumnHeader.Index - 1
    
    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
    ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
On Error GoTo solu
Label4.Caption = ListView1.SelectedItem.Text
'listado = File1.ListIndex
MediaPlayer1.FileName = ListView1.SelectedItem.ListSubItems(1).Text
Slider2.Max = MediaPlayer1.Duration
listado = ListView1.SelectedItem.Index
Text1.SetFocus
Exit Sub
solu:
Select Case Err.Number
Case 380
If Label14.Caption = "Sp" Then
MsgBox "Archivo imposible de reproducir o inexistente."
Else
MsgBox "File impossible to reproduce or it doesn´t exist."
End If
Label4.Caption = ""
End Select


End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Key_Enter
On Error GoTo solu
Label4.Caption = ListView1.SelectedItem.Text
'listado = File1.ListIndex
MediaPlayer1.FileName = ListView1.SelectedItem.ListSubItems(1).Text
Slider2.Max = MediaPlayer1.Duration
listado = ListView1.SelectedItem.Index
Text1.SetFocus
Case Key_F5
On Error GoTo solu
Label4.Caption = ListView1.SelectedItem.Text
'listado = File1.ListIndex
MediaPlayer1.FileName = ListView1.SelectedItem.ListSubItems(1).Text
Slider2.Max = MediaPlayer1.Duration
listado = ListView1.SelectedItem.Index
Text1.SetFocus
Case 46
Call DeleteFiles
If ListView1.ListItems.count > 13 Then
    ListView1.ColumnHeaders(1).Width = 3181
Else
    ListView1.ColumnHeaders(1).Width = 3411
End If
End Select
solu:
Select Case Err.Number
Case 380
If Label14.Caption = "Sp" Then
MsgBox "Archivo imposible de reproducir o inexistente."
Else
MsgBox "File impossible to reproduce or it doesn´t exist."
End If
Label4.Caption = ""
End Select
End Sub

Public Sub DeleteFiles()
    Dim C As Integer
    Dim D As Integer
    
    'On Error GoTo ErrHandler
    For C = 1 To ListView1.ListItems.count
    If ListView1.ListItems(C - D).Selected = True Then
    ListView1.ListItems.Remove ListView1.ListItems(C - D).Index
    D = D + 1
        End If
    
        Next
            Exit Sub
ErrHandler:
    MsgBox "There are no more songs on list."

End Sub
Public Sub DeleteFiles2()
    Dim C As Integer
    Dim D As Integer
    
    'On Error GoTo ErrHandler
    For C = 1 To ListView2.ListItems.count
    If ListView2.ListItems(C - D).Selected = True Then
    ListView2.ListItems.Remove ListView2.ListItems(C - D).Index
    D = D + 1
        End If
    
        Next
            Exit Sub
ErrHandler:
    MsgBox "There are no more songs on list."

End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next

    ListView2.SortKey = ColumnHeader.Index - 1
    
    If ListView2.SortOrder = lvwAscending Then
        ListView2.SortOrder = lvwDescending
    Else
        ListView2.SortOrder = lvwAscending
    End If
    ListView2.Sorted = True
End Sub

Private Sub ListView2_DblClick()
On Error GoTo solu
Label5.Caption = ListView2.SelectedItem.Text
'listado = File1.ListIndex
MediaPlayer2.FileName = ListView2.SelectedItem.ListSubItems(1).Text
Slider3.Max = MediaPlayer2.Duration
listado1 = ListView2.SelectedItem.Index
Text2.SetFocus
Exit Sub
solu:
Select Case Err.Number
Case 380
If Label14.Caption = "Sp" Then
MsgBox "Archivo imposible de reproducir o inexistente."
Else
MsgBox "File impossible to reproduce or it doesn´t exist."
End If
Label4.Caption = ""
End Select

End Sub

Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Key_Enter
On Error GoTo solu
Label5.Caption = ListView2.SelectedItem.Text
'listado = File1.ListIndex
MediaPlayer2.FileName = ListView2.SelectedItem.ListSubItems(1).Text
Slider3.Max = MediaPlayer2.Duration
listado1 = ListView2.SelectedItem.Index
Text2.SetFocus
Case Key_F5
On Error GoTo solu
Label5.Caption = ListView2.SelectedItem.Text
'listado = File1.ListIndex
MediaPlayer2.FileName = ListView2.SelectedItem.ListSubItems(1).Text
Slider3.Max = MediaPlayer2.Duration
listado1 = ListView2.SelectedItem.Index
Text2.SetFocus
Case 46
Call DeleteFiles2
If ListView2.ListItems.count > 13 Then
    ListView2.ColumnHeaders(1).Width = 3651
Else
    ListView2.ColumnHeaders(1).Width = 3890
End If
End Select
solu:
Select Case Err.Number
Case 380
If Label14.Caption = "Sp" Then
MsgBox "Archivo imposible de reproducir o inexistente."
Else
MsgBox "File impossible to reproduce or it doesn´t exist."
End If
Label5.Caption = ""
End Select

End Sub

Private Sub Mousetmr_Timer()
Select Case mousedown
Case MouseR
form2.Show
End Select
End Sub

Private Sub repeat_Click()
If repeat.Value = 0 Then Timer4.Enabled = False: Timer5.Enabled = False: cp.Caption = "C.P. Off": cp.ForeColor = &HFF&
If repeat.Value = 1 Then Timer4.Enabled = True: Timer5.Enabled = True: cp.Caption = "C.P. On": cp.ForeColor = &HFF00&
End Sub

Private Sub Slider1_GotFocus()
If Check1.Value = 0 Then Exit Sub
If Check3.Value = 1 Then
If Slider1.Value = 0 Then
MediaPlayer1.Volume = -6000
MediaPlayer2.Volume = 0
Vol1.Value = 3000
Vol2.Value = -3000
mleft = 1
Exit Sub
End If
If Slider1.Value = -3000 Then
MediaPlayer2.Volume = -6000
MediaPlayer1.Volume = 0
Vol1.Value = -3000
Vol2.Value = 3000
mleft = 0
Exit Sub
End If
If mleft = 0 Then
MediaPlayer1.Volume = -Slider1.Value - 3000
If Slider1.Value >= -1500 Then
MediaPlayer2.Volume = 0
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
MediaPlayer2.Volume = Slider1.Value + 1500
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
If mleft = 1 Then
MediaPlayer2.Volume = Slider1.Value
If Slider1.Value <= -1500 Then
MediaPlayer1.Volume = 0
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
MediaPlayer1.Volume = (-Slider1.Value - 3000) + 1500
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
End If
If Slider1.Value = 0 Then
MediaPlayer1.Volume = -6000
MediaPlayer2.Volume = 0
Vol1.Value = 3000
Vol2.Value = -3000
mleft = 1
Exit Sub
End If
If Slider1.Value = -3000 Then
MediaPlayer2.Volume = -6000
MediaPlayer1.Volume = 0
Vol1.Value = -3000
Vol2.Value = 3000
mleft = 0
Exit Sub
End If
MediaPlayer1.Volume = -Slider1.Value - 3000
MediaPlayer2.Volume = Slider1.Value
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
End Sub

Private Sub Slider1_KeyDown(KeyCode As Integer, Shift As Integer)
If Check1.Value = 0 Then Exit Sub
Select Case KeyCode
Case vbKeyHome
MediaPlayer2.Volume = -6000
MediaPlayer1.Volume = 0
Vol1.Value = -3000
Vol2.Value = 3000
Exit Sub
Case vbKeyEnd
MediaPlayer1.Volume = -6000
MediaPlayer2.Volume = 0
Vol1.Value = 3000
Vol2.Value = -3000
Exit Sub
End Select
If Check3.Value = 1 Then
If Slider1.Value = 0 Then
MediaPlayer1.Volume = -6000
MediaPlayer2.Volume = 0
Vol1.Value = 3000
Vol2.Value = -3000
mleft = 1
Exit Sub
End If
If Slider1.Value = -3000 Then
MediaPlayer2.Volume = -6000
MediaPlayer1.Volume = 0
Vol1.Value = -3000
Vol2.Value = 3000
mleft = 0
Exit Sub
End If
If mleft = 0 Then
MediaPlayer1.Volume = -Slider1.Value - 3000
If Slider1.Value >= -1500 Then
MediaPlayer2.Volume = 0
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
MediaPlayer2.Volume = Slider1.Value + 1500
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
If mleft = 1 Then
MediaPlayer2.Volume = Slider1.Value
If Slider1.Value <= -1500 Then
MediaPlayer1.Volume = 0
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
MediaPlayer1.Volume = (-Slider1.Value - 3000) + 1500
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
End If
If Slider1.Value = 0 Then
MediaPlayer1.Volume = -6000
MediaPlayer2.Volume = 0
Vol1.Value = 3000
Vol2.Value = -3000
mleft = 1
Exit Sub
End If
If Slider1.Value = -3000 Then
MediaPlayer2.Volume = -6000
MediaPlayer1.Volume = 0
Vol1.Value = -3000
Vol2.Value = 3000
mleft = 0
Exit Sub
End If
MediaPlayer1.Volume = -Slider1.Value - 3000
MediaPlayer2.Volume = Slider1.Value
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check1.Value = 0 Then Exit Sub
If Check3.Value = 1 Then
If Slider1.Value = 0 Then
MediaPlayer1.Volume = -6000
MediaPlayer2.Volume = 0
Vol1.Value = 3000
Vol2.Value = -3000
mleft = 1
Exit Sub
End If
If Slider1.Value = -3000 Then
MediaPlayer2.Volume = -6000
MediaPlayer1.Volume = 0
Vol1.Value = -3000
Vol2.Value = 3000
mleft = 0
Exit Sub
End If
If mleft = 0 Then
MediaPlayer1.Volume = -Slider1.Value - 3000
If Slider1.Value >= -1500 Then
MediaPlayer2.Volume = 0
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
MediaPlayer2.Volume = Slider1.Value + 1500
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
If mleft = 1 Then
MediaPlayer2.Volume = Slider1.Value
If Slider1.Value <= -1500 Then
MediaPlayer1.Volume = 0
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
MediaPlayer1.Volume = (-Slider1.Value - 3000) + 1500
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
End If
If Slider1.Value = 0 Then
MediaPlayer1.Volume = -6000
MediaPlayer2.Volume = 0
Vol1.Value = 3000
Vol2.Value = -3000
mleft = 1
Exit Sub
End If
If Slider1.Value = -3000 Then
MediaPlayer2.Volume = -6000
MediaPlayer1.Volume = 0
Vol1.Value = -3000
Vol2.Value = 3000
mleft = 0
Exit Sub
End If
MediaPlayer1.Volume = -Slider1.Value - 3000
MediaPlayer2.Volume = Slider1.Value
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
End Sub




Private Sub Slider10_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyHome
MediaPlayer4.Volume = 0
Slider10.Value = -3000
Exit Sub
Case vbKeyEnd
MediaPlayer4.Volume = -6000
Slider10.Value = 3000
Exit Sub
End Select
If Slider10.Value = 0 Then MediaPlayer4.Volume = -6000: Exit Sub
MediaPlayer4.Volume = -Slider10.Value - 3000
End Sub

Private Sub Slider10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Slider10.Value = 0 Then MediaPlayer4.Volume = -6000: Exit Sub
MediaPlayer4.Volume = -Slider10.Value - 3000
End Sub




Private Sub Slider11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If form1.Label14.Caption = "Sp" Then MsgBox "Solo disponible en la versión registrada.": Exit Sub
MsgBox "Just available on registered version."
Exit Sub
If Slider11.Value = 0 Then MediaPlayer1.Volume = -6000: MediaPlayer2.Volume = -6000: Vol1.Value = 0: Vol2.Value = 0: Exit Sub
If MediaPlayer1.Volume < -Slider11.Value - 3000 Then
GoTo v
Else
MediaPlayer1.Volume = -Slider11.Value - 3000
Vol1.Min = Slider11.Value
Vol1.Value = Slider11.Value
End If
v:
If MediaPlayer2.Volume < -Slider11.Value - 3000 Then
Exit Sub
Else
MediaPlayer2.Volume = -Slider11.Value - 3000
Vol2.Min = Slider11.Value
End If
End Sub

Private Sub Slider2_Click()
MediaPlayer1.CurrentPosition = Slider2.Value
End Sub

Private Sub Slider2_KeyDown(KeyCode As Integer, Shift As Integer)
Timer1.Enabled = False
MediaPlayer1.CurrentPosition = Slider2.Value
End Sub

Private Sub Slider2_KeyUp(KeyCode As Integer, Shift As Integer)
Timer1.Enabled = True
End Sub

Private Sub Slider2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
End Sub

Private Sub Slider2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = True
End Sub

Private Sub Slider3_Click()
MediaPlayer2.CurrentPosition = Slider3.Value
End Sub

Private Sub Slider3_KeyDown(KeyCode As Integer, Shift As Integer)
Timet1.Enabled = False
MediaPlayer2.CurrentPosition = Slider3.Value
End Sub

Private Sub Slider3_KeyUp(KeyCode As Integer, Shift As Integer)
Timet1.Enabled = True
End Sub

Private Sub Slider3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timet1.Enabled = False
End Sub

Private Sub Slider3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timet1.Enabled = True
End Sub

Private Sub Slider4_KeyDown(KeyCode As Integer, Shift As Integer)
If Check2.Value = 1 Then Exit Sub
MediaPlayer1.Balance = Slider4.Value
Slider4.ToolTipText = "Balance A: " & Slider4.Value
End Sub

Private Sub Slider4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check2.Value = 1 Then Exit Sub
MediaPlayer1.Balance = Slider4.Value
Slider4.ToolTipText = "Balance A: " & Slider4.Value
End Sub

Private Sub Slider5_KeyDown(KeyCode As Integer, Shift As Integer)
If Check2.Value = 1 Then Exit Sub
MediaPlayer2.Balance = Slider5.Value
Slider5.ToolTipText = "Balance B: " & Slider5.Value
End Sub

Private Sub Slider5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check2.Value = 1 Then Exit Sub
MediaPlayer2.Balance = Slider5.Value
Slider5.ToolTipText = "Balance B: " & Slider5.Value
End Sub

Private Sub Slider6_GotFocus()
If Check2.Value = 0 Then Slider6.ToolTipText = "Effect Balance: " & Slider6.Value: Exit Sub
If Slider6.Value = 0 Then
MediaPlayer1.Balance = 0
MediaPlayer2.Balance = 0
Exit Sub
End If
MediaPlayer1.Balance = Slider6.Value
MediaPlayer2.Balance = -Slider6.Value
Slider4.Value = MediaPlayer1.Balance
Slider5.Value = MediaPlayer2.Balance
Slider6.ToolTipText = "Effect Balance: " & Slider6.Value
End Sub

Private Sub Slider6_KeyDown(KeyCode As Integer, Shift As Integer)
If Check2.Value = 0 Then Slider6.ToolTipText = "Effect Balance: " & Slider6.Value: Exit Sub
If Slider6.Value = 0 Then
MediaPlayer1.Balance = 0
MediaPlayer2.Balance = 0
Exit Sub
End If
MediaPlayer1.Balance = Slider6.Value
MediaPlayer2.Balance = -Slider6.Value
Slider4.Value = MediaPlayer1.Balance
Slider5.Value = MediaPlayer2.Balance
Slider6.ToolTipText = "Effect Balance: " & Slider6.Value
End Sub

Private Sub Slider6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check2.Value = 0 Then Slider6.ToolTipText = "Effect Balance: " & Slider6.Value: Exit Sub
If Slider6.Value = 0 Then
MediaPlayer1.Balance = 0
MediaPlayer2.Balance = 0
Exit Sub
End If
MediaPlayer1.Balance = Slider6.Value
MediaPlayer2.Balance = -Slider6.Value
Slider4.Value = MediaPlayer1.Balance
Slider5.Value = MediaPlayer2.Balance
Slider6.ToolTipText = "Effect Balance: " & Slider6.Value
End Sub

Private Sub Slider7_Click()
MediaPlayer3.CurrentPosition = Slider7.Value
End Sub

Private Sub Slider7_KeyDown(KeyCode As Integer, Shift As Integer)
Timer0.Enabled = False
MediaPlayer3.CurrentPosition = Slider7.Value
End Sub

Private Sub Slider7_KeyUp(KeyCode As Integer, Shift As Integer)
Timer0.Enabled = True
End Sub

Private Sub Slider7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer0.Enabled = False
End Sub
Private Sub cmdSave_Click()
    Dim sName As String
    
    If WaveMidiFileName = "" Then
        sName = "Radio_from_" & CStr(WaveRecordingStartTime) & "_to_" & CStr(WaveRecordingStopTime)
        sName = Replace(sName, ":", "-")
        sName = Replace(sName, " ", "_")
        sName = Replace(sName, "/", "-")
    Else
        sName = WaveMidiFileName
        sName = Replace(sName, "MID", "wav")
    End If
  
    CommonDialogxx.FileName = sName
    CommonDialogxx.CancelError = True
    On Error GoTo ErrHandler1
    CommonDialogxx.Filter = "WAV file (*.wav*)|*.wav"
    CommonDialogxx.Flags = &H2 Or &H400
    CommonDialogxx.ShowSave
    sName = CommonDialogxx.FileName
    
    WaveSaveAs (sName)
    Exit Sub
ErrHandler1:
End Sub

Private Sub cmdRecord_Click()
'If Form1.Label14.Caption = "Sp" Then MsgBox "Solo disponible en la versión registrada.": Exit Sub
'MsgBox "Just available on registered version."
'Exit Sub
    Dim settings As String
    Dim Alignment As Integer
      
    Alignment = Channels * Resolution / 8
    
    settings = "set capture alignment " & CStr(Alignment) & " bitspersample " & CStr(Resolution) & " samplespersec " & CStr(Rate) & " channels " & CStr(Channels) & " bytespersec " & CStr(Alignment * Rate)
    WaveReset
    WaveSet
    WaveRecord
    WaveRecordingStartTime = Now
    cmdStop.Enabled = True   'Enable the STOP BUTTON
    cmdPlay.Enabled = False  'Disable the "PLAY" button
    cmdSave.Enabled = False  'Disable the "SAVE AS" button
    cmdRecord.Enabled = False 'Disable the "RECORD" button
End Sub

Private Sub cmdSettings_Click()
'If Form1.Label14.Caption = "Sp" Then MsgBox "Solo disponible en la versión registrada.": Exit Sub
'MsgBox "Just available on registered version."
'Exit Sub
Dim strWhat As String
    ' show the user entry form modally
    If form1.Label14.Caption = "Sp" Then strWhat = MsgBox("Si continua, perderá todos los datos guardados en memoria!", vbOKCancel): GoTo english
    strWhat = MsgBox("If you continue your data will be lost!", vbOKCancel)
english:
    If strWhat = vbCancel Then
        Exit Sub
        
        End If
    Sliderxx.Max = 10
    Sliderxx.Value = 0
    Sliderxx.Refresh
    cmdRecord.Enabled = True
    cmdStop.Enabled = False
    cmdPlay.Enabled = False
    cmdSave.Enabled = False
    
    WaveReset
    
    Rate = CLng(GetSetting("MGC DJ 2000", "StartUp", "Rate", "110025"))
    Channels = CInt(GetSetting("MGC DJ 2000", "StartUp", "Channels", "1"))
    Resolution = CInt(GetSetting("MGC DJ 2000", "StartUp", "Resolution", "16"))
    WaveFileName = GetSetting("MGC DJ 2000", "StartUp", "WaveFileName", "C:\Radio.wav")
    WaveAutomaticSave = GetSetting("MGC DJ 2000", "StartUp", "WaveAutomaticSave", "True")

    WaveRecordingImmediate = True
    WaveRecordingReady = False
    WaveRecording = False
    WavePlaying = False
    
    'Be sure to change the Value property of the appropriate button!!
    'if you change the default values!
    
    WaveSet
    frmSettings.optRecordImmediate.Value = True
    frmSettings.Show vbModal
End Sub

Private Sub cmdStop_Click()
    WaveStop
    cmdSave.Enabled = True  'Enable the "SAVE AS" button
    cmdPlay.Enabled = True  'Enable the "PLAY" button
    cmdStop.Enabled = False 'Disable the "STOP" button
    If WavePosition = 0 Then
        Sliderxx.Max = 10
    Else
        If WaveRecordingImmediate And (Not WavePlaying) Then Sliderxx.Max = WavePosition
        If (Not WaveRecordingImmediate) And WaveRecording Then Sliderxx.Max = WavePosition
    End If
    If WaveRecording Then WaveRecordingReady = True
    WaveRecordingStopTime = Now
    WaveRecording = False
    WavePlaying = False
    frmSettings.optRecordProgrammed.Value = False
    frmSettings.optRecordImmediate.Value = True
    frmSettings.lblTimes.Visible = False
End Sub

Private Sub cmdPlay_Click()
    WavePlayFrom (Sliderxx.Value)
    WavePlaying = True
    cmdStop.Enabled = True
    cmdPlay.Enabled = False
End Sub







Private Sub cmdReset_Click()
    Sliderxx.Max = 10
    Sliderxx.Value = 0
    Sliderxx.Refresh
    cmdRecord.Enabled = True
    cmdStop.Enabled = False
    cmdPlay.Enabled = False
    cmdSave.Enabled = False
    
    WaveReset
    
    Rate = CLng(GetSetting("MGC DJ 2000", "StartUp", "Rate", "110025"))
    Channels = CInt(GetSetting("MGC DJ 2000", "StartUp", "Channels", "1"))
    Resolution = CInt(GetSetting("MGC DJ 2000", "StartUp", "Resolution", "16"))
    WaveFileName = GetSetting("MGC DJ 2000", "StartUp", "WaveFileName", "C:\Radio.wav")
    WaveAutomaticSave = GetSetting("MGC DJ 2000", "StartUp", "WaveAutomaticSave", "True")

    WaveRecordingImmediate = True
    WaveRecordingReady = False
    WaveRecording = False
    WavePlaying = False
    WaveMidiFileName = ""
    'Be sure to change the Value property of the appropriate button!!
    'if you change the default values!
    
    WaveSet
    If WaveRenameNecessary Then
        Name WaveShortFileName As WaveLongFileName
        WaveRenameNecessary = False
        WaveShortFileName = ""
    End If
End Sub




Private Sub Sliderxx_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label40.ForeColor = &H0&
Label41.ForeColor = &H0&
Label42.ForeColor = &H0&
Label43.ForeColor = &H0&
Label44.ForeColor = &H0&
Label45.ForeColor = &H0&
End Sub

Private Sub Timer7_Timer()
tempo2 = tempo2 + 1
If tempo2 = 350 Then
If form1.Label14.Caption = "Sp" Then
MediaPlayer1.Stop
MediaPlayer2.Stop
MediaPlayer3.Stop
MediaPlayer4.Stop
MsgBox "Su tiempo de prueba ha finalizado. Para obtener una copia con todas las opciones disponibles y sin límite de tiempo de uso, por favor, siga los pasos que se encuentran en el archivo de ayuda en su idioma y registre su copia de MGC DJ 2000 Gold. Para más información, escriba un mail a register@mgcproductions.com.ar"
Unload form1
End If
MediaPlayer1.Stop
MediaPlayer2.Stop
MediaPlayer3.Stop
MediaPlayer4.Stop
MsgBox "Your trial time has finished, please read the help file and register your copy of MGC DJ 2000 Gold. For more information mail to register@mgcproductions.com.ar": Unload form1
End If
End Sub

Private Sub timerxx_Timer()
    Dim RecordingTimes As String
    Dim msg As String
    
    RecordingTimes = "Start time:  " & WaveRecordingStartTime & vbCrLf _
                    & "Stop time:  " & WaveRecordingStopTime
    
    WaveStatistics
    If Not WaveRecordingImmediate Then
        WaveStatisticsMsg = WaveStatisticsMsg & "Programmed recording"
        If WaveAutomaticSave Then
            WaveStatisticsMsg = WaveStatisticsMsg & " (automatic save)"
        Else
            WaveStatisticsMsg = WaveStatisticsMsg & " (manual save)"
        End If
        WaveStatisticsMsg = WaveStatisticsMsg & vbCrLf & vbCrLf & RecordingTimes
    End If
    StatisticsLabel.Caption = WaveStatisticsMsg
    
    WaveStatus
    If WaveStatusMsg <> form1.Caption Then form1.Caption = WaveStatusMsg
    If InStr(form1.Caption, "stopped") > 0 Then
        Label42.Enabled = False
        Label43.Enabled = True
    End If
    
    If RecordingTimes <> frmSettings.lblTimes.Caption Then frmSettings.lblTimes.Caption = RecordingTimes
    
    If (Now > WaveRecordingStartTime) _
            And (Not WaveRecordingReady) _
            And (Not WaveRecordingImmediate) _
            And (Not WaveRecording) Then
        WaveReset
        WaveSet
        WaveRecord
        WaveRecording = True
        Label42.Enabled = True   'Enable the STOP BUTTON
        Label43.Enabled = False  'Disable the "PLAY" button
        Label44.Enabled = False  'Disable the "SAVE AS" button
        Label41.Enabled = False 'Disable the "RECORD" button
    End If
    
    If (Now > WaveRecordingStopTime) And (Not WaveRecordingReady) And (Not WaveRecordingImmediate) Then
        WaveStop
        Label44.Enabled = True 'Enable the "SAVE AS" button
        Label43.Enabled = True 'Enable the "PLAY" button
        Label42.Enabled = False 'Disable the "STOP" button
        If WavePosition > 0 Then
            Sliderxx.Max = WavePosition
        Else
            Sliderxx.Max = 10
        End If
        WaveRecording = False
        WaveRecordingReady = True
        If WaveAutomaticSave Then
            WaveFileName = "Radio_from_" & CStr(WaveRecordingStartTime) & "_to_" & CStr(WaveRecordingStopTime)
            WaveFileName = Replace(WaveFileName, ":", ".")
            WaveFileName = Replace(WaveFileName, " ", "_")
            WaveFileName = WaveFileName & ".wav"
            WaveSaveAs (WaveFileName)
            msg = "Recording has been saved" & vbCrLf
            msg = msg & "Filename: " & WaveFileName
            MsgBox (msg)
        Else
            msg = "Recording is ready" & vbCrLf
            msg = msg & "Don't forget to save recording..."
            MsgBox (msg)
        End If
        frmSettings.optRecordProgrammed.Value = False
        frmSettings.optRecordImmediate.Value = True
    End If

End Sub

Private Sub Slider7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer0.Enabled = True
End Sub

Private Sub Slider8_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyHome
MediaPlayer3.Volume = 0
Slider8.Value = -3000
Exit Sub
Case vbKeyEnd
MediaPlayer3.Volume = -6000
Slider8.Value = 3000
Exit Sub
End Select
If Slider8.Value = 0 Then MediaPlayer3.Volume = -6000: Exit Sub
MediaPlayer3.Volume = -Slider8.Value - 3000
End Sub

Private Sub Slider8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Slider8.Value = 0 Then MediaPlayer3.Volume = -6000: Exit Sub
MediaPlayer3.Volume = -Slider8.Value - 3000
End Sub
Private Sub Slider9_KeyDown(KeyCode As Integer, Shift As Integer)
MediaPlayer3.Balance = Slider9.Value
Slider9.ToolTipText = "Balance Special: " & Slider9.Value
End Sub

Private Sub Slider9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MediaPlayer3.Balance = Slider9.Value
Slider9.ToolTipText = "Balance Special: " & Slider9.Value
End Sub

Private Sub Text1_Change()
' Crea una variable ListItem.
Dim itmX As ListItem
' Establece la variable al elemento encontrado.
ListView1.MultiSelect = False
Set ListView1.SelectedItem = ListView1.FindItem(Text1.Text, , , lvwPartial)
If ListView1.SelectedItem Is Nothing Then Exit Sub
ListView1.SelectedItem.EnsureVisible
ListView1.MultiSelect = True

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo solu
Select Case KeyCode
Case Key_F5
Label4.Caption = ListView1.SelectedItem.Text
'listado = File1.ListIndex
MediaPlayer1.FileName = ListView1.SelectedItem.ListSubItems(1).Text
Slider2.Max = MediaPlayer1.Duration
listado = ListView1.SelectedItem.Index
Text1.SetFocus
Case Key_Up
ListView1.SetFocus
Case Key_Down
ListView1.SetFocus
Case Key_AvPag
ListView1.SetFocus
Case Key_RePag
ListView1.SetFocus
End Select
Exit Sub
solu:
Select Case Err.Number
Case 380
If Label14.Caption = "Sp" Then
MsgBox "Formato imposible de reproducir."
Else
MsgBox "File impossible to reproduce."
End If
Label4.Caption = ""
Exit Sub
End Select
End Sub

Private Sub Text10_LostFocus()
On Error GoTo mgcgo
If Text10.Text = "" Then Text10.Text = "50"
If Text10.Text < 1 Then Text10.Text = "1"
Exit Sub
mgcgo:
Text10.Text = "50"
End Sub

Private Sub Text2_Change()
' Crea una variable ListItem.
Dim itmX As ListItem
' Establece la variable al elemento encontrado.
ListView2.MultiSelect = False
Set ListView2.SelectedItem = ListView2.FindItem(Text2.Text, , , lvwPartial)
If ListView2.SelectedItem Is Nothing Then Exit Sub
ListView2.SelectedItem.EnsureVisible
ListView2.MultiSelect = True

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo solu
Select Case KeyCode
Case Key_F5
Label5.Caption = ListView2.SelectedItem.Text
'listado = File1.ListIndex
MediaPlayer2.FileName = ListView2.SelectedItem.ListSubItems(1).Text
Slider3.Max = MediaPlayer2.Duration
listado1 = ListView2.SelectedItem.Index
Text2.SetFocus
Case Key_Up
ListView2.SetFocus
Case Key_Down
ListView2.SetFocus
Case Key_AvPag
ListView2.SetFocus
Case Key_RePag
ListView2.SetFocus
End Select
Exit Sub
solu:
Select Case Err.Number
Case 380
If Label14.Caption = "Sp" Then
MsgBox "Formato imposible de reproducir."
Else
MsgBox "File impossible to reproduce."
End If
Label5.Caption = ""
Exit Sub
End Select
End Sub

Private Sub Text3_Change()
Dim Index As Integer
On Error Resume Next
Index = FindFirstMatch(File3, Text3.Text, -1, 0)
File3.ListIndex = Index
If Index >= 0 Then File3.Selected(Index) = True
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo solu
Select Case KeyCode
Case Key_F5
dirc = File3.Path
If Len(File3.Path) > 3 Then dirc = File3.Path & "\"
Label8.Caption = File3.FileName
listado2 = File3.ListIndex
MediaPlayer3.FileName = dirc & File3.FileName
Slider7.Max = MediaPlayer3.Duration
Text3.SetFocus
Case Key_Up
File3.SetFocus
Case Key_Down
File3.SetFocus
Case Key_AvPag
File3.SetFocus
Case Key_RePag
File3.SetFocus
End Select
Exit Sub
solu:
Select Case Err.Number
Case 380
If Label14.Caption = "Sp" Then
MsgBox "Formato imposible de reproducir."
Else
MsgBox "File impossible to reproduce."
End If
Label8.Caption = ""
Exit Sub
End Select
End Sub

Private Sub Text4_LostFocus()
On Error GoTo mgcgo
If Text4.Text = "" Then Text4.Text = "0"
If Text4.Text > 59 Then Text4.Text = "0"
Exit Sub
mgcgo:
Text4.Text = "0"
End Sub

Private Sub Text5_LostFocus()
On Error GoTo mgcgo
If Text5.Text = "" Then Text5.Text = "0"
If Text5.Text > 59 Then Text5.Text = "0"
Exit Sub
mgcgo:
Text5.Text = "0"
End Sub

Private Sub Text6_LostFocus()
On Error GoTo mgcgo
If Text6.Text = "" Then Text6.Text = "0"
If Text6.Text > 59 Then Text6.Text = "0"
Exit Sub
mgcgo:
Text6.Text = "0"
End Sub

Private Sub Text7_LostFocus()
On Error GoTo mgcgo
If Text7.Text = "" Then Text7.Text = "0"
If Text7.Text > 59 Then Text7.Text = "0"
Exit Sub
mgcgo:
Text7.Text = "0"
End Sub

Private Sub Text8_LostFocus()
On Error GoTo mgcgo
If Text8.Text = "" Then Text8.Text = "0"
If Text8.Text > 59 Then Text8.Text = "0"
Exit Sub
mgcgo:
Text8.Text = "0"
End Sub

Private Sub Text9_LostFocus()
On Error GoTo mgcgo
If Text9.Text = "" Then Text9.Text = "0"
If Text9.Text > 59 Then Text9.Text = "0"
Exit Sub
mgcgo:
Text9.Text = "0"
End Sub

Private Sub Timer0_Timer()
Slider7.Value = MediaPlayer3.CurrentPosition
End Sub

Private Sub Timer1_Timer()
Slider2.Value = MediaPlayer1.CurrentPosition
End Sub
Private Sub Timer2_Timer()
Label1.Caption = Time
End Sub

Private Sub Timer3_Timer()
Dim MGC2 As String
Dim ton As String
Dim dag As String
dag = RTrim(tete)
If Timer3.Interval <> 400 Then Timer3.Interval = 400
If durac = Len(dag) Then durac = 0: Timer3.Interval = 2000: Exit Sub
durac = durac + 1
MGC2 = Mid(dag, 1, durac)
MGC.Caption = MGC2
End Sub

Private Sub Timer4_Timer()
Dim i As Integer
On Error GoTo solu
Dim chero As Integer
For i = 1 To ListView1.ListItems.count
If ListView1.ListItems(i).Index = 2 Then GoTo ok1
Next i
Exit Sub
ok1:
If listado = ListView1.ListItems.count Then
MediaPlayer1.FileName = ListView1.ListItems(listado).ListSubItems(1).Text
Label4.Caption = ListView1.ListItems(listado).Text
Slider2.Max = MediaPlayer1.Duration
listado = 1
Exit Sub
End If
If listado < 0 Then Exit Sub
If MediaPlayer1.CurrentPosition = MediaPlayer1.Duration Then
listado = listado + 1
MediaPlayer1.FileName = ListView1.ListItems(listado).ListSubItems(1).Text
Slider2.Max = MediaPlayer1.Duration
Label4.Caption = ListView1.ListItems(listado).Text
End If
Exit Sub
solu:
Select Case Err.Number
Case 380
MediaPlayer1.Stop
Exit Sub
End Select
End Sub

Private Sub Timer5_Timer()
On Error GoTo solu
Dim chero As Integer
Dim i As Integer
For i = 1 To ListView2.ListItems.count
If ListView2.ListItems(i).Index = 2 Then GoTo ok1
Next i
Exit Sub
ok1:
If listado1 = ListView2.ListItems.count Then
MediaPlayer2.FileName = ListView2.ListItems(listado1).ListSubItems(1).Text
Label5.Caption = ListView2.ListItems(listado1).Text
Slider3.Max = MediaPlayer2.Duration
listado1 = 1
Exit Sub
End If
If listado1 < 0 Then Exit Sub
If MediaPlayer2.CurrentPosition = MediaPlayer2.Duration Then
listado1 = listado1 + 1
MediaPlayer2.FileName = ListView2.ListItems(listado1).ListSubItems(1).Text
Slider3.Max = MediaPlayer2.Duration
Label5.Caption = ListView2.ListItems(listado1).Text
End If
Exit Sub
solu:
Select Case Err.Number
Case 380
MediaPlayer2.Stop
Exit Sub
End Select
End Sub

Private Sub Timer6_Timer()
If Check1.Value = 0 Then GoTo mgc33
If Direction = "null" Then GoTo mgc33
If Direction = "right" Then
If Slider1.Value = 0 Then Direction = "null": MediaPlayer1.Volume = -6000: MediaPlayer2.Volume = 0: Vol1.Value = 3000: Vol2.Value = -3000: mleft = 1: Exit Sub
Slider1.Value = Slider1.Value + 50
If Check3.Value = 1 Then
If mleft = 0 Then
MediaPlayer1.Volume = -Slider1.Value - 3000
If Slider1.Value >= -1500 Then
MediaPlayer2.Volume = 0
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
MediaPlayer2.Volume = Slider1.Value + 1500
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
If mleft = 1 Then
MediaPlayer2.Volume = Slider1.Value
If Slider1.Value <= -1500 Then
MediaPlayer1.Volume = 0
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
MediaPlayer1.Volume = (-Slider1.Value - 3000) + 1500
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
End If
MediaPlayer1.Volume = -Slider1.Value - 3000
MediaPlayer2.Volume = Slider1.Value
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
End If
If Direction = "left" Then
If Slider1.Value = -3000 Then Direction = "null": MediaPlayer2.Volume = -6000: MediaPlayer1.Volume = 0: Vol1.Value = -3000: Vol2.Value = 3000: mleft = 0: Exit Sub
Slider1.Value = Slider1.Value - 50
If Check3.Value = 1 Then
If mleft = 0 Then
MediaPlayer1.Volume = -Slider1.Value - 3000
If Slider1.Value >= -1500 Then
MediaPlayer2.Volume = 0
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
MediaPlayer2.Volume = Slider1.Value + 1500
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
If mleft = 1 Then
MediaPlayer2.Volume = Slider1.Value
If Slider1.Value <= -1500 Then
MediaPlayer1.Volume = 0
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
MediaPlayer1.Volume = (-Slider1.Value - 3000) + 1500
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
Exit Sub
End If
End If
MediaPlayer1.Volume = -Slider1.Value - 3000
MediaPlayer2.Volume = Slider1.Value
Vol1.Value = -MediaPlayer1.Volume - 3000
Vol2.Value = -MediaPlayer2.Volume - 3000
End If
Exit Sub
mgc33:
Timer6.Enabled = False
End Sub

Private Sub Timet1_Timer()
Slider3.Value = MediaPlayer2.CurrentPosition
End Sub

Private Sub Tmr2Time_Timer()
Dim Pri1 As Single
Dim Pri11 As Single
Dim terc1 As Single
Dim tec11 As Single
Dim seg1 As Integer
Dim seg11 As Integer
Dim Cuart1 As Integer
Dim Cuart11 As Integer
Dim tela1 As String
Dim tela11 As String
Pri1 = MediaPlayer2.Duration / 60
seg1 = Pri1
terc1 = (((Pri1 - seg1) * 100) * 60) / 100
Cuart1 = terc1
If MediaPlayer2.FileName = "" Then Label3.Caption = "0:00 / 0:00": Exit Sub
If Cuart1 < 0 Then seg1 = seg1 - 1: Cuart1 = 60 + Cuart1
Pri11 = MediaPlayer2.CurrentPosition / 60
seg11 = Pri11
terc11 = (((Pri11 - seg11) * 100) * 60) / 100
Cuart11 = terc11
If Cuart11 < 0 Then seg11 = seg11 - 1: Cuart11 = 60 + Cuart11
tela1 = Cuart1
tela11 = Cuart11
If Cuart11 < 10 Then tela11 = "0" & Cuart11
If Cuart1 < 10 Then tela1 = "0" & Cuart1
Label3.Caption = seg11 & ":" & tela11 & " / " & seg1 & ":" & tela1
End Sub

Private Sub TmrTime_Timer()
Dim Pri As Single
Dim Pri1 As Single
Dim terc As Single
Dim tec1 As Single
Dim seg As Integer
Dim seg1 As Integer
Dim Cuart As Integer
Dim Cuart1 As Integer
Dim tela As String
Dim tela1 As String
Pri = MediaPlayer1.Duration / 60
seg = Pri
terc = (((Pri - seg) * 100) * 60) / 100
Cuart = terc
If MediaPlayer1.FileName = "" Then Label2.Caption = "0:00 / 0:00": Exit Sub
If Cuart < 0 Then seg = seg - 1: Cuart = 60 + Cuart
Pri1 = MediaPlayer1.CurrentPosition / 60
seg1 = Pri1
terc1 = (((Pri1 - seg1) * 100) * 60) / 100
Cuart1 = terc1
If Cuart1 < 0 Then seg1 = seg1 - 1: Cuart1 = 60 + Cuart1
tela = Cuart
tela1 = Cuart1
If Cuart1 < 10 Then tela1 = "0" & Cuart1
If Cuart < 10 Then tela = "0" & Cuart
Label2.Caption = seg1 & ":" & tela1 & " / " & seg & ":" & tela
End Sub

Private Sub tupdate_Timer()
If form1.Label14.Caption = "Sp" Then
Dim Version2 As String, News2 As String, Dir2 As String
    On Error GoTo ErrorMessage2
    'now assign content of file application.ver to variable Version
    Version2 = Inet1.OpenURL("http://orbita.starmedia.com/~mgcproductions/update/applications.ver")
    'You can try this function on Your local disk, but You must change adresses:
    'for example: "file://c:\path\application.ver"
    If Version2 = "" Then GoTo ErrorMessage 'if file not found or file is empty then exit
    If Version2 <= App.Major & "." & App.Minor Then
                GoTo ErrorMessage
    End If
    'now display MessageBox with news in newer version(s) of application and two buttons Yes(update), No(end)
    News2 = Inet1.OpenURL("http://orbita.starmedia.com/~mgcproductions/update/newss.txt")
       If MsgBox(News2 & Version2, vbYesNo, "Actualize su versión " & App.Major & "." & App.Minor & " a la nueva versión " & Version2) = vbYes Then
    Dir2 = Inet1.OpenURL("http://orbita.starmedia.com/~mgcproductions/update/applications.dir")
        HyperJump Dir2 'this will run default download manager (probable also open default browser)
    GoTo ErrorMessage
    End If
MsgBox "Puede actualizar su version actual de Mgc Dj 2000 manualmente en http://www.mgcproductions.com.ar"
ErrorMessage2:
tupdate.Enabled = False
Exit Sub
End If
'This function assume files "application.ver", "news.txt" and "application.zip"
'on server http://server.com/user (change "server.com/user" by your server name and path)
'Inspect contain of files "news.txt" and "application.ver" at examples
Dim Version As String, News As String, Dir As String
    On Error GoTo ErrorMessage
    'now assign content of file application.ver to variable Version
    Version = Inet1.OpenURL("http://orbita.starmedia.com/~mgcproductions/update/application.ver")
    'You can try this function on Your local disk, but You must change adresses:
    'for example: "file://c:\path\application.ver"
    If Version = "" Then GoTo ErrorMessage 'if file not found or file is empty then exit
    If Version <= App.Major & "." & App.Minor Then
                GoTo ErrorMessage
    End If
    'now display MessageBox with news in newer version(s) of application and two buttons Yes(update), No(end)
    News = Inet1.OpenURL("http://orbita.starmedia.com/~mgcproductions/update/news.txt")
       If MsgBox(News & Version, vbYesNo, "You can update from version " & App.Major & "." & App.Minor & " to version " & Version) = vbYes Then
    Dir = Inet1.OpenURL("http://orbita.starmedia.com/~mgcproductions/update/application.dir")
        HyperJump Dir 'this will run default download manager (probable also open default browser)
    GoTo ErrorMessage
    End If
MsgBox "You can download new version of Mgc Dj 2000 manually at http://www.mgcproductions.com.ar"
ErrorMessage:
tupdate.Enabled = False
End Sub

Private Sub Vol1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyHome
MediaPlayer1.Volume = 0
Vol1.Value = -3000
Exit Sub
Case vbKeyEnd
MediaPlayer1.Volume = -6000
Vol1.Value = 3000
Exit Sub
End Select
If Vol1.Value = 0 Then MediaPlayer1.Volume = -6000: Exit Sub
MediaPlayer1.Volume = -Vol1.Value - 3000
End Sub

Private Sub Vol1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check1.Value = 1 Then Exit Sub
If Vol1.Value = 0 Then MediaPlayer1.Volume = -6000: Exit Sub
MediaPlayer1.Volume = -Vol1.Value - 3000
End Sub

Private Sub Vol2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyHome
MediaPlayer2.Volume = 0
Vol2.Value = -3000
Exit Sub
Case vbKeyEnd
MediaPlayer2.Volume = -6000
Vol2.Value = 3000
Exit Sub
End Select
If Vol2.Value = 0 Then MediaPlayer2.Volume = -6000: Exit Sub
MediaPlayer2.Volume = -Vol2.Value - 3000
End Sub

Private Sub Vol2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check1.Value = 1 Then Exit Sub
If Vol2.Value = 0 Then MediaPlayer2.Volume = -6000: Exit Sub
MediaPlayer2.Volume = -Vol2.Value - 3000
End Sub
