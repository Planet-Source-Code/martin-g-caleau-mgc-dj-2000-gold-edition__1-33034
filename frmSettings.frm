VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "MGC DJ Wav Settings"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSettings.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdMidi 
      Caption         =   "&Karaoke: choose a midi file to record"
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      ToolTipText     =   "Then you select a midi file, press ´OK´ and then ´Record´"
      Top             =   3240
      Width           =   3615
   End
   Begin VB.CommandButton cmdOke 
      Caption         =   "&OK"
      Height          =   1095
      Left            =   4560
      TabIndex        =   13
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HELP"
      Height          =   1095
      Left            =   4560
      TabIndex        =   15
      ToolTipText     =   "Help to record mixes, dj´s effects, mic, midi, etc."
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2175
      Left            =   1200
      ScaleHeight     =   2115
      ScaleWidth      =   3315
      TabIndex        =   16
      Top             =   1080
      Width           =   3375
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Sample rate (Hz)"
         ForeColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton optRate44100 
            BackColor       =   &H80000012&
            Caption         =   "44100"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   360
            TabIndex        =   27
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optRate22050 
            BackColor       =   &H80000012&
            Caption         =   "22050"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   360
            TabIndex        =   26
            Top             =   660
            Width           =   1095
         End
         Begin VB.OptionButton optRate11025 
            BackColor       =   &H80000012&
            Caption         =   "11025"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   360
            TabIndex        =   25
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton optRate8000 
            BackColor       =   &H80000012&
            Caption         =   "8000"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   360
            TabIndex        =   24
            Top             =   1260
            Width           =   1095
         End
         Begin VB.OptionButton optRate6000 
            BackColor       =   &H80000012&
            Caption         =   "6000"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   360
            TabIndex        =   23
            Top             =   1560
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Channels"
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   1680
         TabIndex        =   20
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton optStereo 
            BackColor       =   &H00000000&
            Caption         =   "stereo"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optMono 
            BackColor       =   &H00000000&
            Caption         =   "mono"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "Resolution"
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   1680
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
         Begin VB.OptionButton opt16bits 
            BackColor       =   &H00000000&
            Caption         =   "16 bits"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   360
            TabIndex        =   19
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton opt8bits 
            BackColor       =   &H00000000&
            Caption         =   "8 bits"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   480
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame6 
      Caption         =   "Recoding options"
      Height          =   135
      Left            =   6480
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   135
      Begin VB.OptionButton optRecordImmediate 
         Caption         =   "Manual recording"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optRecordProgrammed 
         Caption         =   "Programmed recording"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   2055
      End
      Begin VB.Frame frmTimes 
         Caption         =   "Enter times"
         Height          =   1575
         Left            =   2520
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton cmdStartTime 
            Caption         =   "Start time"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdStopTime 
            Caption         =   "Stop time"
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   960
            Width           =   1695
         End
      End
      Begin VB.Frame frmManualAuto 
         Caption         =   "Saving file"
         Height          =   1695
         Left            =   2520
         TabIndex        =   3
         Top             =   1680
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton cmdFileName 
            Caption         =   "Filename"
            Height          =   375
            Left            =   360
            TabIndex        =   12
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Manual"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Automatic"
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   720
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Label lblTimes 
         Caption         =   " "
         Height          =   1215
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Record Setting"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit

Private Sub cmdFileName_Click()
    WaveFileName = InputBox("Filename: ", "Filename for automatic saving", WaveFileName)
End Sub

Private Sub cmdMidi_Click()
    CommonDialog2.CancelError = True
    On Error GoTo ErrHandler1
    CommonDialog2.Filter = "Midi file (*.mid*)|*.mid"
    CommonDialog2.Flags = &H2 Or &H400
    CommonDialog2.ShowOpen
    WaveMidiFileName = CommonDialog2.FileName
    WaveMidiFileName = GetShortName(WaveMidiFileName)
ErrHandler1:
End Sub

Private Sub cmdOke_Click()
    Unload Me
End Sub

Private Sub cmdStartTime_Click()
    Dim wrst As String
    wrst = WaveRecordingStartTime
    wrst = InputBox("Enter start time recording", "Start time", wrst)
    If wrst = "" Then Exit Sub
    If Not IsDate(wrst) Then
        MsgBox ("The date/time you entered was not valid!")
    Else
    ' String returned from InputBox is a valid time,
    ' so store it as a date/time value in WaveRecordingStartTime.
        If CDate(wrst) < Now Then
            MsgBox ("Recording events in the past is not possible...")
            WaveRecordingStartTime = Now + TimeSerial(0, 15, 0)
        Else
            WaveRecordingStartTime = CDate(wrst)
        End If
        If WaveRecordingStopTime < WaveRecordingStartTime Then WaveRecordingStopTime = WaveRecordingStartTime + TimeSerial(0, 15, 0)
    End If
End Sub

Private Sub cmdStopTime_Click()
    Dim wrst As String
    
    wrst = WaveRecordingStopTime
    If wrst < WaveRecordingStartTime Then wrst = WaveRecordingStartTime + TimeSerial(0, 15, 0)
        
    wrst = InputBox("Enter stop time recording", "Stop time", wrst)
    If wrst = "" Then Exit Sub
    If Not IsDate(wrst) Then
        MsgBox ("The time you entered was not valid!")
    Else
    ' String returned from InputBox is a valid time,
    ' so store it as a date/time value in WaveRecordingStartTime.
        If CDate(wrst) < WaveRecordingStartTime Then
            MsgBox ("The stop time has to be later then the start time!")
            WaveRecordingStopTime = WaveRecordingStartTime + TimeSerial(0, 5, 0)
        Else
            WaveRecordingStopTime = CDate(wrst)
        End If
    End If
End Sub

Private Sub Command1_Click()
Dim Error1 As Integer
On Error Resume Next
If form1.Label14.Caption = "Sp" Then Error1 = ShellExecute(Me.hWnd, "Open", App.Path & "\HELP\Spanish\mgcdj2000faqs\Record_Mix\record_mix.html", "", "", 1): Exit Sub
Error1 = ShellExecute(Me.hWnd, "Open", App.Path & "\HELP\English\mgcdj2000faq\Record_Mix\record_mix.html", "", "", 1)
End Sub

Private Sub Form_Deactivate()
Unload frmSettings
End Sub

Private Sub Form_Load()
'On Error GoTo errend

Dim rgn1 As Long    'main region
    Dim rgn2 As Long    'region to combine with rgn1
    Dim rc As Long        'return code or looping index
    rgn1 = CreateEllipticRgn(400, 300, 1, 1)           'create region 1
    rc = SetWindowRgn(Me.hWnd, rgn1, True)
Move form1.Left + 2700, form1.Top + 2500
If form1.Label14.Caption = "Sp" Then
cmdMidi.Caption = "Karaoke: seleccione un archivo midi para grabar"
cmdMidi.ToolTipText = "Luego de seleccionar el archivo midi, presione ´OK´ y posteriormente ´Grabar´."
Command1.Caption = "AYUDA"
Command1.ToolTipText = "Ayuda para grabar mixes, efectos dj, voces, etc."
End If
If form1.Label6.Caption <> "" Then
    frmSettings.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorfsettings)
End If

    Select Case Rate
    Case 44100
        optRate44100.Value = True
    Case 22050
        optRate22050.Value = True
    Case 11025
        optRate11025.Value = True
    Case 8000
        optRate8000.Value = True
    Case 6000
        optRate6000.Value = True
    End Select
    
    Select Case Channels
    Case 1
        optMono.Value = True
    Case 2
        optStereo.Value = True
    End Select
    
    Select Case Resolution
    Case 8
        opt8bits.Value = True
    Case 16
        opt16bits.Value = True
    End Select
    
    'If WaveRecordingImmediate Then
    '   optRecordImmediate.Value = True
    'else
    '    optRecordProgrammed.Value = True
   ' End If
   
    'If WaveAutomaticSave Then
    '    Option11.Value = True
    'Else
    '    Option10.Value = True
    'End If
errend:
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call SaveSetting("MGC DJ 2000", "StartUp", "Rate", CStr(Rate))
    Call SaveSetting("MGC DJ 2000", "StartUp", "Channels", CStr(Channels))
    Call SaveSetting("MGC DJ 2000", "StartUp", "Resolution", CStr(Resolution))
    'Call SaveSetting("MGC DJ 2000", "StartUp", "WaveFileName", WaveFileName)
    'Call SaveSetting("MGC DJ 2000", "StartUp", "WaveAutomaticSave", CStr(WaveAutomaticSave))
    WaveReset
    
    'WaveRecordingImmediate = True
    WaveRecordingReady = False
    WaveRecording = False
    WavePlaying = False
    
    'Be sure to change the Value property of the appropriate button!!
    'if you change the default values!
    
    WaveSet
    'WaveRecordingStartTime = Now + TimeSerial(0, 15, 0)
    'WaveRecordingStopTime = WaveRecordingStartTime + TimeSerial(0, 15, 0)
    'WaveMidiFileName = ""
    WaveRenameNecessary = False
End Sub

Private Sub optRate11025_Click()
    Rate = 11025
    optRate11025.Value = True
End Sub

Private Sub optRate44100_Click()
    Rate = 44100
    optRate44100.Value = True
End Sub

Private Sub Option10_Click()
    WaveAutomaticSave = False
End Sub

Private Sub Option11_Click()
    WaveAutomaticSave = True
End Sub

Private Sub optRate22050_Click()
    Rate = 22050
    optRate22050.Value = True
End Sub


Private Sub optRate8000_Click()
    Rate = 8000
    optRate8000.Value = True
End Sub

Private Sub optRate6000_Click()
    Rate = 6000
    optRate6000.Value = True
End Sub

Private Sub optMono_Click()
    Channels = 1
    optMono.Value = True
End Sub

Private Sub optStereo_Click()
    Channels = 2
    optStereo.Value = True
End Sub

Private Sub opt8bits_Click()
    Resolution = 8
    opt8bits.Value = True
End Sub

Private Sub opt16bits_Click()
    Resolution = 16
    opt16bits.Value = True
End Sub

Private Sub optRecordImmediate_Click()
    WaveRecordingImmediate = True
    frmManualAuto.Visible = False
    frmTimes.Visible = False
    lblTimes.Visible = False
    form1.Label41.Enabled = True
End Sub

Private Sub optRecordProgrammed_Click()
    WaveRecordingImmediate = False
    frmManualAuto.Visible = True
    frmTimes.Visible = True
    lblTimes.Visible = True
    form1.Label41.Enabled = False
    If WaveRecordingStartTime < Now Then
        WaveRecordingStartTime = Now + TimeSerial(0, 15, 0)
        WaveRecordingStopTime = WaveRecordingStartTime + TimeSerial(0, 15, 0)
    End If

End Sub

