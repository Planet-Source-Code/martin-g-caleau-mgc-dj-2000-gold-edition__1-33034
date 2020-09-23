VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mgcmixer 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "MGC DJ Mixer"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4965
   Icon            =   "mgcmixer.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "mgcmixer.frx":0442
   ScaleHeight     =   2310
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Stay on Top"
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
      Height          =   135
      Left            =   600
      TabIndex        =   16
      Top             =   480
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1720
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   20
      SmallChange     =   4
      Min             =   -100
      Max             =   0
      SelectRange     =   -1  'True
      SelStart        =   -100
      SelLength       =   100
      TickFrequency   =   5
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   975
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1720
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   20
      SmallChange     =   4
      Min             =   -100
      Max             =   0
      SelectRange     =   -1  'True
      SelStart        =   -100
      SelLength       =   100
      TickFrequency   =   5
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   975
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1720
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   20
      SmallChange     =   4
      Min             =   -100
      Max             =   0
      SelectRange     =   -1  'True
      SelStart        =   -100
      SelLength       =   100
      TickFrequency   =   5
   End
   Begin MSComctlLib.Slider Slider4 
      Height          =   975
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1720
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   20
      SmallChange     =   4
      Min             =   -100
      Max             =   0
      SelectRange     =   -1  'True
      SelStart        =   -100
      SelLength       =   100
      TickFrequency   =   5
   End
   Begin MSComctlLib.Slider Slider5 
      Height          =   975
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1720
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   20
      SmallChange     =   4
      Min             =   -100
      Max             =   0
      SelectRange     =   -1  'True
      SelStart        =   -100
      SelLength       =   100
      TickFrequency   =   5
   End
   Begin MSComctlLib.Slider Slider6 
      Height          =   975
      Left            =   3600
      TabIndex        =   5
      Top             =   960
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1720
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   20
      SmallChange     =   4
      Min             =   -100
      Max             =   0
      SelectRange     =   -1  'True
      SelStart        =   -100
      SelLength       =   100
      TickFrequency   =   5
   End
   Begin MSComctlLib.Slider Slider7 
      Height          =   975
      Left            =   4200
      TabIndex        =   6
      Top             =   960
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1720
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   20
      SmallChange     =   4
      Min             =   -100
      Max             =   0
      SelectRange     =   -1  'True
      SelStart        =   -100
      SelLength       =   100
      TickFrequency   =   5
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   4170
      TabIndex        =   15
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "MGC DJ MIXER"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Spk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4150
      TabIndex        =   13
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Mic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3570
      TabIndex        =   12
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Line In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2800
      TabIndex        =   11
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2370
      TabIndex        =   10
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Midi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1720
      TabIndex        =   9
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Wave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1050
      TabIndex        =   8
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   550
      TabIndex        =   7
      Top             =   720
      Width           =   375
   End
End
Attribute VB_Name = "mgcmixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Const NO_BUTTON = 0
Const WM_NCLBUTTONDOWN = &HA1
Dim t1 As String
Dim sl1 As Long
Dim mousedown As Integer



Private Sub Check1_Click()
If Check1.Value = 1 Then
SetWindowPos hWnd, conHwndTopmost, 235, 50, 332, 154, conSwpNoActivate Or conSwpShowWindow
Move form1.Left + 3350, form1.Top + 400
Exit Sub
End If
SetWindowPos hWnd, 1, 235, 50, 332, 154, conSwpNoActivate Or conSwpShowWindow
Move form1.Left + 3350, form1.Top + 400
End Sub

Private Sub Form_Load()
If form1.Label6.Caption <> "" Then
    mgcmixer.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colormixer)
End If
SetWindowPos hWnd, conHwndTopmost, 235, 50, 332, 154, conSwpNoActivate Or conSwpShowWindow
Move form1.Left + 3350, form1.Top + 400
Dim v1 As Long
Dim v2 As Long
Dim v3 As Long
Dim v4 As Long
Dim v5 As Long
Dim v6 As Long
Dim v7 As Long
w = lGetVolume(v1, v1, 5)
w = lGetVolume(v2, v2, 0)
w = lGetVolume(v3, v3, 1)
w = lGetVolume(v4, v4, 2)
w = lGetVolume(v5, v5, 3)
w = lGetVolume(v6, v6, 4)
w = lGetVolume(v7, v7, 6)
Slider1.Value = -v1 * 100 / 12
Slider2.Value = -v2 * 100 / 12
Slider3.Value = -v3 * 100 / 12
Slider4.Value = -v4 * 100 / 12
Slider5.Value = -v5 * 100 / 12
Slider6.Value = -v6 * 100 / 12
Slider7.Value = -v7 * 100 / 12
If form1.Label14.Caption = "Sp" Then
Check1.Caption = "Posición Top"
Label9.Caption = "Salir"
End If
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
Label9.ForeColor = &HFFFFFF
If mgcmixer.Top > form1.Top + 100 And mgcmixer.Top < form1.Top + 700 Then Movetest = 0: Exit Sub
Movetest = 1
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mousedown = NO_BUTTON
End Sub

Private Sub Form_Unload(Cancel As Integer)
form1.Check4.Value = 0
End Sub

Private Sub Label9_Click()
Unload mgcmixer
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = &HFF&
End Sub

Private Sub Slider1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyHome
z = lSetVolume(12, 12, 5)
Slider1.Value = -100
Exit Sub
Case vbKeyEnd
z = lSetVolume(0, 0, 5)
Slider1.Value = 0
Exit Sub
End Select
z = lSetVolume(-Slider1.Value * 12 / 100, -Slider1.Value * 12 / 100, 5)
End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "si"
End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider1.Text = Str(-Slider1.Value)
Slider1.ToolTipText = Slider1.Text
If t1 = "si" Then z = lSetVolume(-Slider1.Value * 12 / 100, -Slider1.Value * 12 / 100, 5): Exit Sub
Dim LeftC As Long
w = lGetVolume(LeftC, LeftC, 5)
Slider1.Value = -LeftC * 100 / 12
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "no"
End Sub

Private Sub Slider2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyHome
z = lSetVolume(12, 12, 0)
Slider2.Value = -100
Exit Sub
Case vbKeyEnd
z = lSetVolume(0, 0, 0)
Slider2.Value = 0
Exit Sub
End Select
z = lSetVolume(-Slider2.Value * 12 / 100, -Slider2.Value * 12 / 100, 0)
End Sub

Private Sub Slider2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "si"
End Sub

Private Sub Slider2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider2.Text = Str(-Slider2.Value)
Slider2.ToolTipText = Slider2.Text
If t1 = "si" Then z = lSetVolume(-Slider2.Value * 12 / 100, -Slider2.Value * 12 / 100, 0): Exit Sub
Dim LeftC As Long
w = lGetVolume(LeftC, LeftC, 0)
Slider2.Value = -LeftC * 100 / 12
End Sub

Private Sub Slider2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "no"
End Sub

Private Sub Slider3_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyHome
z = lSetVolume(12, 12, 1)
Slider3.Value = -100
Exit Sub
Case vbKeyEnd
z = lSetVolume(0, 0, 1)
Slider3.Value = 0
Exit Sub
End Select
z = lSetVolume(-Slider3.Value * 12 / 100, -Slider3.Value * 12 / 100, 1)
End Sub

Private Sub Slider3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "si"
End Sub

Private Sub Slider3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider3.Text = Str(-Slider3.Value)
Slider3.ToolTipText = Slider3.Text
If t1 = "si" Then z = lSetVolume(-Slider3.Value * 12 / 100, -Slider3.Value * 12 / 100, 1): Exit Sub
Dim LeftC As Long
w = lGetVolume(LeftC, LeftC, 1)
Slider3.Value = -LeftC * 100 / 12
End Sub

Private Sub Slider3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "no"
End Sub

Private Sub Slider4_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyHome
z = lSetVolume(12, 12, 2)
Slider4.Value = -100
Exit Sub
Case vbKeyEnd
z = lSetVolume(0, 0, 2)
Slider4.Value = 0
Exit Sub
End Select
z = lSetVolume(-Slider4.Value * 12 / 100, -Slider4.Value * 12 / 100, 2)
End Sub

Private Sub Slider4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "si"
End Sub

Private Sub Slider4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider4.Text = Str(-Slider4.Value)
Slider4.ToolTipText = Slider4.Text
If t1 = "si" Then z = lSetVolume(-Slider4.Value * 12 / 100, -Slider4.Value * 12 / 100, 2): Exit Sub
Dim LeftC As Long
w = lGetVolume(LeftC, LeftC, 2)
Slider4.Value = -LeftC * 100 / 12
End Sub

Private Sub Slider4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "no"
End Sub

Private Sub Slider5_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyHome
z = lSetVolume(12, 12, 3)
Slider5.Value = -100
Exit Sub
Case vbKeyEnd
z = lSetVolume(0, 0, 3)
Slider5.Value = 0
Exit Sub
End Select
z = lSetVolume(-Slider5.Value * 12 / 100, -Slider5.Value * 12 / 100, 3)
End Sub

Private Sub Slider5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "si"
End Sub

Private Sub Slider5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider5.Text = Str(-Slider5.Value)
Slider5.ToolTipText = Slider5.Text
If t1 = "si" Then z = lSetVolume(-Slider5.Value * 12 / 100, -Slider5.Value * 12 / 100, 3): Exit Sub
Dim LeftC As Long
w = lGetVolume(LeftC, LeftC, 3)
Slider5.Value = -LeftC * 100 / 12
End Sub

Private Sub Slider5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "no"
End Sub

Private Sub Slider6_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyHome
z = lSetVolume(12, 12, 4)
Slider6.Value = -100
Exit Sub
Case vbKeyEnd
z = lSetVolume(0, 0, 4)
Slider6.Value = 0
Exit Sub
End Select
z = lSetVolume(-Slider6.Value * 12 / 100, -Slider6.Value * 12 / 100, 4)
End Sub

Private Sub Slider6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "si"
End Sub

Private Sub Slider6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider6.Text = Str(-Slider6.Value)
Slider6.ToolTipText = Slider6.Text
If t1 = "si" Then z = lSetVolume(-Slider6.Value * 12 / 100, -Slider6.Value * 12 / 100, 4): Exit Sub
Dim LeftC As Long
w = lGetVolume(LeftC, LeftC, 4)
Slider6.Value = -LeftC * 100 / 12
End Sub

Private Sub Slider6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "no"
End Sub

Private Sub Slider7_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyHome
z = lSetVolume(12, 12, 6)
Slider7.Value = -100
Exit Sub
Case vbKeyEnd
z = lSetVolume(0, 0, 6)
Slider7.Value = 0
Exit Sub
End Select
z = lSetVolume(-Slider7.Value * 12 / 100, -Slider7.Value * 12 / 100, 6)
End Sub

Private Sub Slider7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "si"
End Sub

Private Sub Slider7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider7.Text = Str(-Slider7.Value)
Slider7.ToolTipText = Slider7.Text
If t1 = "si" Then z = lSetVolume(-Slider7.Value * 12 / 100, -Slider7.Value * 12 / 100, 6): Exit Sub
Dim LeftC As Long
w = lGetVolume(LeftC, LeftC, 6)
Slider7.Value = -LeftC * 100 / 12
End Sub

Private Sub Slider7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
t1 = "no"
End Sub
