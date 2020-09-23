VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   5925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":000C
   ScaleHeight     =   5925
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   3015
      Left            =   840
      ScaleHeight     =   2955
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   1440
      Width           =   5535
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   2415
         Left            =   2040
         ScaleHeight     =   2415
         ScaleWidth      =   3495
         TabIndex        =   1
         Top             =   240
         Width           =   3495
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   7215
            Left            =   0
            ScaleHeight     =   7215
            ScaleWidth      =   3255
            TabIndex        =   2
            Top             =   840
            Width           =   3255
            Begin VB.Frame Frame2 
               BackColor       =   &H00FF0000&
               Height          =   15
               Left            =   240
               TabIndex        =   4
               Top             =   3960
               Width           =   2895
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00FF0000&
               Height          =   15
               Left            =   240
               TabIndex        =   3
               Top             =   2520
               Width           =   2895
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Caption         =   "Martín Gustavo Caleau"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   1680
               Width           =   2895
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackColor       =   &H8000000E&
               BackStyle       =   0  'Transparent
               Caption         =   "Free v4.1 - Build: 25/03/2002"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   240
               TabIndex        =   12
               Top             =   480
               Width           =   2655
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackColor       =   &H80000012&
               Caption         =   "~ 2002 ~"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   960
               TabIndex        =   11
               Top             =   6960
               Width           =   1215
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackColor       =   &H80000008&
               Caption         =   "República Argentina"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFC0&
               Height          =   255
               Left            =   720
               TabIndex        =   10
               Top             =   6720
               Width           =   1695
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H8000000E&
               BackStyle       =   0  'Transparent
               Caption         =   "MGC DJ 2000 Gold"
               BeginProperty Font 
                  Name            =   "Impact"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   495
               Left            =   240
               TabIndex        =   9
               Top             =   0
               Width           =   2655
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Caption         =   "by MGC PRODUCTIONS(c)"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   1320
               Width           =   2895
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Caption         =   "www.mgcproductions.com.ar"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   255
               Left            =   240
               TabIndex        =   7
               Top             =   840
               Width           =   2655
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               Caption         =   "This special version was made in memory of RODRIGO ""El Potro de Córdoba""."
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   735
               Left            =   120
               TabIndex        =   6
               Top             =   3000
               Width           =   3015
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               Caption         =   $"Form3.frx":2B94
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
               Height          =   1815
               Left            =   120
               TabIndex        =   5
               Top             =   4080
               Width           =   3015
            End
            Begin VB.Image Image1 
               Height          =   975
               Left            =   720
               Picture         =   "Form3.frx":2C58
               Stretch         =   -1  'True
               Top             =   5760
               Width           =   1650
            End
         End
      End
      Begin VB.Image Image3 
         Height          =   735
         Left            =   360
         Picture         =   "Form3.frx":3D91
         ToolTipText     =   "Made in ARGENTINA"
         Top             =   240
         Width           =   1050
      End
      Begin VB.Image Image2 
         Height          =   1575
         Left            =   240
         Picture         =   "Form3.frx":4ECA
         ToolTipText     =   "Author: Martín G. Caleau / www.mgcproductions.com.ar"
         Top             =   1080
         Width           =   1410
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Form_Deactivate()
Unload form3
End Sub

Private Sub Form_Load()
On Error GoTo fin
If form1.Label6.Caption <> "" Then
    form3.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf3)
End If
Move form1.Left + 2200, form1.Top + 1000
Dim rgn1 As Long    'main region
    Dim rgn2 As Long    'region to combine with rgn1
    Dim rc As Long        'return code or looping index
    rgn1 = CreateEllipticRgn(475, 390, 1, 1)           'create region 1
    rc = SetWindowRgn(Me.hWnd, rgn1, True)
If form1.Label14.Caption = "Sp" Then
form3.Label4 = "Esta versión especial ha sido realizada en memoria de Rodrigo 'El Potro de Córdoba' en el día de su fallecimiento."
form3.Label5 = "Espero que el programa cumpla con sus objetivos: Lograr una fácil interacción usuario-computadora y al mismo tiempo experimentada, para que puedas trabajar con la música como un verdadero DJ. Muy útil para Estaciones de Radio y Djs."
Image2.ToolTipText = "Logo de MGC PRODUCTIONS / www.mgcproductions.com.ar"
End If
Exit Sub
fin:
MsgBox "Skin not founded. You have to Re-Install MGC DJ 2000."
form1.SetFocus
Unload form1
Unload form3
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF00&
End Sub

Private Sub Image2_Click()
Dim Error As Integer
On Error Resume Next
Error = ShellExecute(Me.hWnd, "Open", "www.mgcproductions.com.ar", "", "", 1)
End Sub

Private Sub Label3_Click()
Dim Error As Integer
On Error Resume Next
Error = ShellExecute(Me.hWnd, "Open", "www.mgcproductions.com.ar", "", "", 1)
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFFFFFF
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF00&
End Sub

Private Sub Timer1_Timer()
If Picture4.Top = -7200 Then Picture4.Top = 2700
Picture4.Top = Picture4.Top - 10
End Sub
