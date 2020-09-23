VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   LinkTopic       =   "Form5"
   Picture         =   "Formrodrigo.frx":0000
   ScaleHeight     =   6015
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   9435
      TabIndex        =   2
      Top             =   4800
      Width           =   9495
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "This special version was made in memory of RODRIGO ""El Potro de Córdoba"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   2280
         TabIndex        =   3
         Top             =   120
         Width           =   5055
      End
   End
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   6000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Formrodrigo.frx":1CFB
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Interval        =   520
      Left            =   480
      Top             =   600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EL REY DEL CUARTETO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3360
      Left            =   480
      Picture         =   "Formrodrigo.frx":2641
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   5055
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
If form1.Label14.Caption = "Sp" Then
Label3.Caption = "Esta versión especial fue realizada en memoria de Rodrigo ´El Potro de Córdoba´"
End If
End Sub

Private Sub Form_LostFocus()
Unload Form5
End Sub

Private Sub Timer1_Timer()
If Label2.Visible = True Then
Label2.Visible = False
Else
Label2.Visible = True
End If
End Sub
