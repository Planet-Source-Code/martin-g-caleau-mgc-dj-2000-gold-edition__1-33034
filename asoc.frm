VERSION 5.00
Begin VB.Form asociate 
   Caption         =   "Form7"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "asociate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function fCreateShellLink Lib "VB5STKIT.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long

Private Sub Form_Load()
    Dim strString As String
    Dim lngDword As Long

If Command$ = "" Then GoTo yeta:
            MsgBox (Command$ & " is the file you need To open!"), vbInformation
        'Add to Recent file folder
        'lReturn = fCreateShellLink("..\..\Recent", Command$, Command$, "")
    
    'create an entry in the class key
    
Exit Sub
yeta:
MsgBox "acabas de abrir mgc dj 2000"
Call savestring(HKEY_CLASSES_ROOT, "\.mp3", "", "MGC DJ 2000 Music File")
    'content type
    Call savestring(HKEY_CLASSES_ROOT, "\.mp3", "Content Type", "audio/x-wav")
    'name
    Call savestring(HKEY_CLASSES_ROOT, "\MGC DJ 2000 Music File", "", "MGC DJ 2000 Music File")
    'edit flags
    Call SaveDword(HKEY_CLASSES_ROOT, "\MGC DJ 2000 Music File", "EditFlags", "0000")
    'file's icon (can be an icon file, or an
    '     icon located within a dll file)
    Call savestring(HKEY_CLASSES_ROOT, "\MGC DJ 2000 Music File\DefaultIcon", "", App.Path & "MGC DJ 2000 Gold.exe")
    'Shell
    Call savestring(HKEY_CLASSES_ROOT, "\MGC DJ 2000 Music File\Shell", "", "")
    'Shell Open
    Call savestring(HKEY_CLASSES_ROOT, "\MGC DJ 2000 Music File\Shell\Open", "", "")
    'Shell open command
    Call savestring(HKEY_CLASSES_ROOT, "\MGC DJ 2000 Music File\Shell\Open\command", "", App.Path & "MGC DJ 2000 Gold.exe %1")
End Sub


