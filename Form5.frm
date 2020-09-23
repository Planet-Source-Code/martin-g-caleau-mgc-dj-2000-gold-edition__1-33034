VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   6735
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3855
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form5.frx":28EE
      Top             =   1560
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3855
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form5.frx":4329
      Top             =   1560
      Width           =   4695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Deactivate()
Unload Form5
End Sub

Private Sub Form_Load()
On Error GoTo fin
If Form1.Label6.Caption = "yellow" Then Form5.Picture = LoadPicture("skins\help-yellow.jpg")
If Form1.Label6.Caption = "red   " Then Form5.Picture = LoadPicture("skins\help-red.jpg")
If Form1.Label6.Caption = "blue  " Then Form5.Picture = LoadPicture("skins\help-blue.jpg")
If Form1.Label6.Caption = "sa    " Then Form5.Picture = LoadPicture("skins\help-sp1.jpg")
If Form1.Label6.Caption = "sb    " Then Form5.Picture = LoadPicture("skins\help-sp2.jpg")
Dim rgn1 As Long    'main region
    Dim rgn2 As Long    'region to combine with rgn1
    Dim rc As Long        'return code or looping index
    rgn1 = CreateEllipticRgn(450, 450, 1, 1)           'create region 1
    rc = SetWindowRgn(Me.hWnd, rgn1, True)
Move Form1.Left + 2300, Form1.Top + 700
'Call ExplodeForm(Me, 400)
If Form1.Label14.Caption = "Sp" Then
Text1.Visible = False
Text2.Visible = True
End If
Exit Sub
fin:
MsgBox "Skin not founded. You have to Re-Install MGC DJ 2000."
Unload Form5
Unload Form1
End Sub

