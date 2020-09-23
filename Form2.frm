VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":000C
   ScaleHeight     =   4770
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rodrigo ""El Potro"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DJ Mixer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dj´s Effects Setup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Customize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Help / Advertising"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   3360
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Form_Load()
On Error GoTo fin

If form1.Label6.Caption <> "" Then
    form2.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf2)
End If

Move form1.Left + 3100, form1.Top + 2300
Dim rgn1 As Long    'main region
    Dim rgn2 As Long    'region to combine with rgn1
    Dim rc As Long        'return code or looping index
    rgn1 = CreateEllipticRgn(375, 315, 1, 1)           'create region 1
    rc = SetWindowRgn(Me.hWnd, rgn1, True)
If form1.Label14.Caption = "Sp" Then

form2.Label1.Caption = "Configurar"
form2.Label2.Caption = "Salir"
form2.Label3.Caption = "About"
form2.Label4.Caption = "Ayuda / Publicidad"
form2.Label5.Caption = "Efectos Dj Setup"
End If
Exit Sub
fin:
MsgBox "Skin not found."
form1.Label6.Caption = ""
form1.SetFocus
Unload form1
Unload form2
End Sub

Private Sub Form_LostFocus()
Unload form2
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &H80000007
Label2.ForeColor = &H80000007
Label4.ForeColor = &H80000007
Label1.ForeColor = &H80000007
Label3.ForeColor = &H80000007
Label6.ForeColor = &H80000007
Label7.ForeColor = &H80000007
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &H80000007
Label2.ForeColor = &H80000007
Label4.ForeColor = &H80000007
Label1.ForeColor = &H80000007
Label3.ForeColor = &H80000007
Label6.ForeColor = &H80000007
Label7.ForeColor = &H80000007
End Sub

Private Sub Label1_Click()
Form4.Show
Form4.SetFocus
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF&
Label2.ForeColor = &H80000007
Label4.ForeColor = &H80000007
Label5.ForeColor = &H80000007
Label3.ForeColor = &H80000007
Label6.ForeColor = &H80000007
Label7.ForeColor = &H80000007
End Sub

Private Sub Label2_Click()
Unload form2
Unload form1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF&
Label4.ForeColor = &H80000007
Label1.ForeColor = &H80000007
Label3.ForeColor = &H80000007
Label5.ForeColor = &H80000007
Label6.ForeColor = &H80000007
Label7.ForeColor = &H80000007
End Sub

Private Sub Label3_Click()
form3.Show
form3.SetFocus
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF&
Label2.ForeColor = &H80000007
Label4.ForeColor = &H80000007
Label1.ForeColor = &H80000007
Label5.ForeColor = &H80000007
Label6.ForeColor = &H80000007
Label7.ForeColor = &H80000007
End Sub

Private Sub Label4_Click()
Dim Error As Integer
On Error Resume Next
If form1.Label14.Caption = "Sp" Then Error = ShellExecute(Me.hWnd, "Open", App.Path & "\Help\indexspanish.htm", "", "", 1): Exit Sub
Error = ShellExecute(Me.hWnd, "Open", App.Path & "\Help\index.htm", "", "", 1)
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF&
Label2.ForeColor = &H80000007
Label1.ForeColor = &H80000007
Label3.ForeColor = &H80000007
Label5.ForeColor = &H80000007
Label6.ForeColor = &H80000007
Label7.ForeColor = &H80000007
End Sub

Private Sub Label5_Click()
Form6.Show
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &H80000007
Label2.ForeColor = &H80000007
Label4.ForeColor = &H80000007
Label5.ForeColor = &HFF&
Label3.ForeColor = &H80000007
Label6.ForeColor = &H80000007
Label7.ForeColor = &H80000007
End Sub

Private Sub Label6_Click()
form1.Check4.Value = 1
mgcmixer.Show
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &H80000007
Label2.ForeColor = &H80000007
Label4.ForeColor = &H80000007
Label5.ForeColor = &H80000007
Label3.ForeColor = &H80000007
Label6.ForeColor = &HFF&
Label7.ForeColor = &H80000007
End Sub

Private Sub Label7_Click()
Form5.Show
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &H80000007
Label4.ForeColor = &H80000007
Label1.ForeColor = &H80000007
Label3.ForeColor = &H80000007
Label5.ForeColor = &H80000007
Label6.ForeColor = &H80000007
Label7.ForeColor = &HFF&
End Sub


