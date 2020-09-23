VERSION 5.00
Begin VB.Form resolut 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MGC SCREEN RESOLUTION ERROR"
   ClientHeight    =   4665
   ClientLeft      =   1125
   ClientTop       =   1500
   ClientWidth     =   6465
   Icon            =   "Changrez.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4665
   ScaleWidth      =   6465
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   4035
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "&CHANGE to Selected Resolution"
         Height          =   615
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1455
         Left            =   4440
         Picture         =   "Changrez.frx":27A2
         Stretch         =   -1  'True
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Available Resolutions (Resoluciones disponibles)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"Changrez.frx":47D2
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"Changrez.frx":4867
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   3480
         Width           =   5775
      End
   End
End
Attribute VB_Name = "resolut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' ChangRez sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
'
' How to change video resolution in Windows 95.
' The WinSDK API declarations for VB does NOT include
' this useful procedure, nor does it include some
' of the needed constants.  I had to guess at the API
' declaration (well, an educated guess ;-), and go to
' my Visual C include files to find the DM_ constants.

Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpInitData As DEVMODE, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As DEVMODE) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

' /* Flags for ChangeDisplaySettings */
Const CDS_UPDATEREGISTRY = &H1
Const CDS_TEST = &H2
Const CDS_FULLSCREEN = &H4
Const CDS_GLOBAL = &H8
Const CDS_SET_PRIMARY = &H10
Const CDS_RESET = &H40000000
Const CDS_SETRECT = &H20000000
Const CDS_NORESET = &H10000000

' /* Return values for ChangeDisplaySettings */
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
Const DISP_CHANGE_FAILED = -1
Const DISP_CHANGE_BADMODE = -2
Const DISP_CHANGE_NOTUPDATED = -3
Const DISP_CHANGE_BADFLAGS = -4
Const DISP_CHANGE_BADPARAM = -5

Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4

Dim D() As DEVMODE, lNumModes As Long

Private Sub Command1_Click()
Dim twidth As Integer
    Dim l As Long, Flags As Long, X As Long
    X = List1.ListIndex + 1
    D(X).dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
    Flags = CDS_UPDATEREGISTRY
    l = ChangeDisplaySettings(D(X), Flags)
    Select Case l
        Case DISP_CHANGE_RESTART
            l = MsgBox("You must reboot for the change to take effect.", vbOKCancel)
            If l = vbOK Then
                Flags = 0
                l = ExitWindowsEx(EWX_REBOOT, Flags)
            End If
        Case DISP_CHANGE_SUCCESSFUL
        Case Else
            MsgBox "Error changing resolution!"
    End Select
twidth = Screen.Width \ Screen.TwipsPerPixelX
If twidth < 800 Then MsgBox "Screen Resolution Error. You must select a 800x600 or high resolution. (Error de resolución de pantalla. Debes cambiar la resolución a 800x600 o superior)."
If twidth >= 800 Then MsgBox "The error was repaired with success !!!.(Error Solucionado !!!).": Load form1: form1.Show: Unload resolut
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
Move (Screen.Width - resolut.Width) / 2, (Screen.Height - resolut.Height) / 2
    Dim l As Long, lMaxModes As Long
    lMaxModes = 8
    ReDim D(0 To lMaxModes) As DEVMODE
    lNumModes = 0
    l = EnumDisplaySettings(0, lNumModes, D(lNumModes))
    Do
        lNumModes = lNumModes + 1
        If lNumModes > lMaxModes Then
            lMaxModes = lMaxModes + 8
            ReDim Preserve D(0 To lMaxModes) As DEVMODE
        End If
        l = EnumDisplaySettings(0, lNumModes, D(lNumModes))
        If l = 0 Then Exit Do
        List1.AddItem D(lNumModes).dmPelsWidth & "x" & D(lNumModes).dmPelsHeight & "x" & D(lNumModes).dmBitsPerPel
    Loop
    lNumModes = lNumModes - 1
Dim twidth As Integer
twidth = Screen.Width \ Screen.TwipsPerPixelX
If twidth >= 800 Then form1.Show: Unload resolut
End Sub

