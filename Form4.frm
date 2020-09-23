VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
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
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   7005
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1440
      Left            =   3900
      Pattern         =   "*.skn"
      TabIndex        =   10
      ToolTipText     =   "Double Click: Change to selected Skin / Download more skins from: www.mgcproductions.com.ar"
      Top             =   1200
      Width           =   2655
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   9600
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   225
      Left            =   7440
      MaxLength       =   13
      TabIndex        =   1
      Text            =   "Martín Caleau"
      ToolTipText     =   "Change to desired dj text and then press ´DJ TEXT´"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.OptionButton Option8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "English"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9600
      TabIndex        =   4
      ToolTipText     =   "English Languaje"
      Top             =   480
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton Option9 
      BackColor       =   &H00000000&
      Caption         =   "Español"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   8640
      TabIndex        =   3
      ToolTipText     =   "Idioma Español"
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox mgcsound 
      BackColor       =   &H00000000&
      Caption         =   "INI SOUND ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   8640
      TabIndex        =   2
      ToolTipText     =   "MGC DJ 2000 START SOUND"
      Top             =   240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DJ TEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   6360
      TabIndex        =   12
      ToolTipText     =   "Put your desired Front Dj Text"
      Top             =   4245
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "English"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      ToolTipText     =   "Clickea aquí para cambiar a idioma Español"
      Top             =   5475
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INI SOUND ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      ToolTipText     =   "Enable or Disable the start sound"
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Recommend Mgc Dj 2000 to a Friend"
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
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Make Mgc Dj 2000 your Default Player"
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
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      ToolTipText     =   "Register to Mgc Dj 2000 as your default player in mp3, wav and midi files"
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Update to latest New Free Version of MGC Dj 2000"
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
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      ToolTipText     =   "http://www.mgcproductions.com.ar"
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Language:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   5475
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select a SKIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4270
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim soundc As String
Dim lencolor As Long

Private Sub Command2_Click()
Form6.Show
End Sub



Private Sub Command4_Click()
Dim Error As Integer
On Error Resume Next
Error = ShellExecute(Me.hWnd, "Open", "www.mgcproductions.com.ar", "", "", 1)
End Sub


Private Sub File1_DblClick()
On Error GoTo solu
form1.Label6.Caption = File1.FileName

If form1.Label6.Caption = "default.skn" Then
    form1.Label6.Caption = ""
    req = 1
    Unload form1
    form1.Show
    form1.SetFocus
Exit Sub
End If

If Label6.Caption <> "" Then
On Error GoTo solu

Dim skiner As Skin
Dim Regskin As Integer
numCanal = FreeFile
Open App.Path & "\skins\" & form1.Label6.Caption For Random As #numCanal Len = 264
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
Unload Form4

Exit Sub
solu:
MsgBox "Background file impossible to find."
End Sub

Private Sub Form_Deactivate()
Unload Form4
End Sub

Private Sub Form_Load()
'On Error GoTo fin
Text1.Text = RTrim(tete)
If form1.Label6.Caption <> "" Then
    Form4.Picture = LoadPicture(App.Path & "/skins/" & colordirname & "\" & colorf4)
End If
Dim rgn1 As Long    'main region
    Dim rgn2 As Long    'region to combine with rgn1
    Dim rc As Long        'return code or looping index

    rgn1 = CreateEllipticRgn(675, 450, 1, 1)           'create region 1
    rc = SetWindowRgn(Me.hWnd, rgn1, True)

Move form1.Left + 700, form1.Top + 900
'File1.Path = App.Path & "\Skins"
lencolor = &HFF&
If form1.Label14.Caption = "Sp" Then
Label1.Caption = "Seleccione un SKIN"
File1.ToolTipText = "Doble Click: Cambia al SKIN seleccionado / Baje más skins desde: www.mgcproductions.com.ar"
Label3.Caption = "Lenguaje"
Label4.Caption = "Actualizar a Nueva Versión Free de MGC Dj 2000"
Label8.Caption = "TEXTO DJ"
Text1.ToolTipText = "Cambie al texto dj deseado y luego presione en ´TEXTO DJ´"
Label6.ToolTipText = "Activa o Desactiva el sonido de inicio"
Option8.Value = False
Option9.Value = True
Label7.Caption = "Español"
Label7.ForeColor = &H80FF&
lencolor = &H80FF&
Label8.ToolTipText = "Coloca el texto DJ del Inicio deseado."
Label5.Caption = "Hacer de MGC DJ 2000 tu DJ Player"
Label2.Caption = "Recomendar Mgc Dj 2000 a un Amigo"
Label7.ToolTipText = "Click Here to change to English language"
Label5.ToolTipText = "Registrar a Mgc Dj 2000 como su reproductor por defecto en archivos mp3, wav y midis"
End If
If form1.Label14.Caption = "En" Then Option8.Value = True: Label7.Caption = "English": Label7.ForeColor = &HFF&: lencolor = &HFF&
If form1.Label10.Caption = "S" Then mgcsound.Value = 1: Label6.Caption = "INI SOUND ON": Label6.ForeColor = &HFF00&
If form1.Label10.Caption = "N" Then mgcsound.Value = 0: Label6.Caption = "INI SOUND OFF": Label6.ForeColor = &HFF&
soundc = Label6.ForeColor
Exit Sub
fin:
MsgBox "Skin not founded. You have to Re-Install MGC DJ 2000."
Unload form1
Unload Form4
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFFFF00
Label5.ForeColor = &HFF00&
Label4.ForeColor = &HFFFF00
Label2.ForeColor = &HFF00&
Label6.ForeColor = soundc
Label7.ForeColor = lencolor
End Sub

Private Sub Label2_Click()
email.Show
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFFFFFF
Label4.ForeColor = &HFFFF00
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Please, check at http://www.mgcproductions.com.ar for updates."
Exit Sub
Label4.ForeColor = &HFFFF&
If IsConnected = True Then
If form1.Label14.Caption = "Sp" Then
Dim Version2 As String, News2 As String, Dir2 As String
    On Error GoTo ErrorMessage2
    Me.MousePointer = 11
    'now assign content of file application.ver to variable Version
    Version2 = Inet1.OpenURL("http://orbita.starmedia.com/~mgcproductions/update/applications.ver")
    'You can try this function on Your local disk, but You must change adresses:
    'for example: "file://c:\path\application.ver"
    If Version2 = "" Then GoTo Skip2 'if file not found or file is empty then exit
    If Version2 <= App.Major & "." & App.Minor Then
        MsgBox "Su versión Actual de Mgc Dj 2000 es la última versión disponible. No ha sido realizada una nueva versión.", vbInformation
        GoTo Skip2
    End If
    'now display MessageBox with news in newer version(s) of application and two buttons Yes(update), No(end)
    News2 = Inet1.OpenURL("http://orbita.starmedia.com/~mgcproductions/update/newss.txt")
       If MsgBox(News2 & Version2, vbYesNo, "Actualize su versión " & App.Major & "." & App.Minor & " a la nueva versión " & Version2) = vbYes Then
    Dir2 = Inet1.OpenURL("http://orbita.starmedia.com/~mgcproductions/update/applications.dir")
        HyperJump Dir2 'this will run default download manager (probable also open default browser)
    Exit Sub
    End If
MsgBox "Puede actualizar su nueva versión manualmente en http://www.mgcproductions.com.ar"
Skip2:
    Me.MousePointer = 0
    Exit Sub
ErrorMessage2:
    Me.MousePointer = 0
    MsgBox "Ha ocurrido un error. Falló la actualización." & Chr(10) & "Debe actualizar su versión de Mgc Dj 2000 manualmente en http://www.mgcproductions.com.ar", vbCritical
Exit Sub
    End If
    'This function assume files "application.ver", "news.txt" and "application.zip"
'on server http://server.com/user (change "server.com/user" by your server name and path)
'Inspect contain of files "news.txt" and "application.ver" at examples
Dim Version As String, News As String, Dir As String
    On Error GoTo ErrorMessage
    Me.MousePointer = 11
    'now assign content of file application.ver to variable Version
    Version = Inet1.OpenURL("http://orbita.starmedia.com/~mgcproductions/update/application.ver")
    'You can try this function on Your local disk, but You must change adresses:
    'for example: "file://c:\path\application.ver"
    If Version = "" Then GoTo Skip 'if file not found or file is empty then exit
    If Version <= App.Major & "." & App.Minor Then
        MsgBox "You have the latest version available of Mgc Dj 2000. No new version was released.", vbInformation
        GoTo Skip
    End If
    'now display MessageBox with news in newer version(s) of application and two buttons Yes(update), No(end)
    News = Inet1.OpenURL("http://orbita.starmedia.com/~mgcproductions/update/news.txt")
       If MsgBox(News & Version, vbYesNo, "You can update from version " & App.Major & "." & App.Minor & " to version " & Version) = vbYes Then
    Dir = Inet1.OpenURL("http://orbita.starmedia.com/~mgcproductions/update/application.dir")
        HyperJump Dir 'this will run default download manager (probable also open default browser)
    Exit Sub
    End If
MsgBox "You can download new version of Mgc Dj 2000 manually at http://www.mgcproductions.com.ar"
Skip:
    Me.MousePointer = 0
    Exit Sub
ErrorMessage:
    Me.MousePointer = 0
    MsgBox "An error has occured. Update failed." & Chr(10) & "You must download new version of Mgc Dj 2000 manually at http://www.mgcproductions.com.ar", vbCritical
Exit Sub
    End If
    If form1.Label14.Caption = "Sp" Then
    MsgBox "Por Favor, conéctese a Internet para chequear por nuevas versiones disponibles de Mgc Dj 2000."
    Exit Sub
    End If
    MsgBox "Please, connect it to Internet to check Latest Version available of Mgc Dj 2000."
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF00&
Label4.ForeColor = &HFFFFFF
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
End Sub

Private Sub Label5_Click()
On Error GoTo fin
 Call savestring(HKEY_CLASSES_ROOT, "\.mp3", "", "MGC DJ 2000 Music File")
    'content type
    Call savestring(HKEY_CLASSES_ROOT, "\.mp3", "Content Type", "audio/x-wav")
    'name
        'edit flags
    Call savestring(HKEY_CLASSES_ROOT, "\.wav", "", "MGC DJ 2000 Music File")
    Call savestring(HKEY_CLASSES_ROOT, "\.wav", "Content Type", "audio/x-wav")
    
    Call savestring(HKEY_CLASSES_ROOT, "\.mid", "", "MGC DJ 2000 Music File")
    Call savestring(HKEY_CLASSES_ROOT, "\.mid", "Content Type", "audio/x-wav")
    
    Call savestring(HKEY_CLASSES_ROOT, "\MGC DJ 2000 Music File", "", "MGC DJ 2000 Music File")
    
    Call SaveDword(HKEY_CLASSES_ROOT, "\MGC DJ 2000 Music File", "EditFlags", "0000")
    'file's icon (can be an icon file, or an
    '     icon located within a dll file)
    Call savestring(HKEY_CLASSES_ROOT, "\MGC DJ 2000 Music File\DefaultIcon", "", App.Path & "\MGC DJ 2000 Gold.exe")
    'Shell
    Call savestring(HKEY_CLASSES_ROOT, "\MGC DJ 2000 Music File\Shell", "", "")
    'Shell Open
    Call savestring(HKEY_CLASSES_ROOT, "\MGC DJ 2000 Music File\Shell\Open", "", "")
    'Shell open command
    Call savestring(HKEY_CLASSES_ROOT, "\MGC DJ 2000 Music File\Shell\Open\command", "", App.Path & "\MGC DJ 2000 Gold.exe %1")
    MsgBox "Associate Files Successful: mp3, wav and midi.", , "MGC DJ 2000 Association"
fin:
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &HFFFFFF
Label7.ForeColor = lencolor
End Sub

Private Sub Label6_Click()
If mgcsound.Value = 1 Then form1.Label10.Caption = "N": Label6.Caption = "INI SOUND OFF": Label6.ForeColor = &HFF&: soundc = &HFF&: mgcsound.Value = 0: Exit Sub
If mgcsound.Value = 0 Then form1.Label10.Caption = "S": Label6.Caption = "INI SOUND ON": Label6.ForeColor = &HFF00&: soundc = &HFF00&: mgcsound.Value = 1
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFFFFFF
End Sub

Private Sub Label7_Click()
If Option8.Value = False Then
form1.Label14.Caption = "En"
req = 1
Unload form1
form1.Show
form1.SetFocus
Option9.Value = False
Option8.Value = True
Label7.Caption = "English"
Label7.Caption = &HFF&
lencolor = &HFF&
Exit Sub
End If
If Option9.Value = False Then
form1.Label14.Caption = "Sp"
req = 1
Unload form1
form1.Show
form1.SetFocus
Option8.Value = False
Option9.Value = True
Label7.Caption = "Español"
Label7.ForeColor = &H80FF&
lencolor = &H80FF&
End If
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFFFFFF
Label5.ForeColor = &HFF00&
End Sub

Private Sub Label8_Click()
Dim fret As String
fret = Trim(Text1.Text)
If fret = "" Then tete = "MGC Dj 2000": form1.Labeldj.Caption = "MGC Dj 2000": form1.SetFocus: Exit Sub
tete = Text1.Text
form1.Labeldj.Caption = Text1.Text
form1.SetFocus
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFFFFFF
End Sub

Private Sub mgcsound_Click()
If mgcsound.Value = 0 Then form1.Label10.Caption = "N": mgcsound.Caption = "INI SOUND OFF": mgcsound.ForeColor = &HFF&
If mgcsound.Value = 1 Then form1.Label10.Caption = "S": mgcsound.Caption = "INI SOUND ON": mgcsound.ForeColor = &HFF00&
End Sub





Private Sub Option8_Click()
Option9.Value = False
End Sub
Private Sub Option9_Click()
Option8.Value = False
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFFFF00
End Sub
