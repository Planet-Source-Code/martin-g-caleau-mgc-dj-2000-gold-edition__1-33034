VERSION 5.00
Begin VB.Form email 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "MGC Friend´s Mail"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "email.frx":0000
   ScaleHeight     =   2310
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   2400
      Width           =   3135
      Begin VB.Label Label2 
         Caption         =   "Verificación:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Todavía NO VERIFICADO"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MGC Dj 2000 will auto-send the E-mail to your Friend recommending the program."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   540
      Width           =   3855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Friend´s Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Friend´s E-mail:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "email"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
Dim texto As String
Dim inimail As String
Dim finmail As String
Dim segok As Boolean
Dim priok As Boolean
Dim i As Integer
Dim X As Integer
Dim count As Integer
    count = 0
    priok = False
    segok = False
    texto = Text1.Text
    If Len(texto) = 0 Then GoTo ver
    
    For i = 1 To Len(texto)
        inimail = Mid(texto, i, 1)
        If inimail = "@" Then count = count + 1
    Next i
    
    If count = 1 Then priok = True
    
    For X = Len(texto) To 1 Step -1
        finmail = Mid(texto, X, 1)
        If finmail = "." Then segok = True
        If Mid(texto, Len(texto), 1) = "." Then segok = False
        If Mid(texto, Len(texto), 1) = "@" Then segok = False
        If Mid(texto, 1, 1) = "." Then segok = False
        If Mid(texto, 1, 1) = "@" Then segok = False
    Next X
ver:
    If priok = True And segok = True Then
        GoTo todook
    Else
        MsgBox "Invalid Information."
        Exit Sub
    End If
todook:
SetStringValue "HKEY_LOCAL_MACHINE\Software\MGC DJ 2000", "FriendName", Text2.Text
SetStringValue "HKEY_LOCAL_MACHINE\Software\MGC DJ 2000", "FriendEmail", Text1.Text
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "DJ Mail", WinDir + "Djmail.exe"
If form1.Label14.Caption = "Sp" Then
    SetStringValue "HKEY_LOCAL_MACHINE\Software\MGC DJ 2000", "Lang", "Sp"
Else
    SetStringValue "HKEY_LOCAL_MACHINE\Software\MGC DJ 2000", "Lang", "En"
End If
Dim error2 As Integer
error2 = ShellExecute(Me.hWnd, "Open", WinDir & "Djmail.exe", "", "", 1)
If IsConnected = True Then
    If form1.Label14.Caption = "Sp" Then
        MsgBox "Enviando el E-Mail de Mgc Dj 2000 a su amigo . . ."
        Unload email
    Else
        MsgBox "Sending the Mgc Dj 2000 E-mail Message to your friend . . ."
        Unload email
    End If
Else
    If form1.Label14.Caption = "Sp" Then
        MsgBox "El E-Mail de Mgc Dj 2000 será enviado a su amigo cuando Ud. se conecte a Internet."
        Unload email
    Else
        MsgBox "The Mgc Dj 2000 E-mail will be sended to your friend when you connect it to Internet."
        Unload email
    End If
End If
    Exit Sub
errorx:
MsgBox "Mail Error. The e-mail hasn´t been sended."
End Sub



Private Sub Form_Load()
If form1.Label6.Caption <> "" Then
    email.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & coloremail)
End If
If form1.Label14.Caption = "Sp" Then
    Label5.Caption = "Mgc Dj 2000 enviará automáticamente el e-mail a su amigo recomendando el programa."
    Label4.Caption = "Nombre de Amigo"
    Label1.Caption = "E-mail de Amigo"
    Command1.Caption = "&Enviar"
    Label6.Caption = "Salir"
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFFFFFF
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFFFFFF
End Sub

Private Sub Label6_Click()
Unload email
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFF&
End Sub
