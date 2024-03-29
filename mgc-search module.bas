Attribute VB_Name = "SearchModule"
Option Explicit
Option Compare Text

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

#If Win32 Then
    Public Const CB_FINDSTRING = &H14C
    Public Const CB_FINDSTRINGEXACT = &H158
    Public Const FB_FINDSTRING = &H18F
    Public Const FB_FINDSTRINGEXACT = &H1A2
#Else
    Public Const WM_USER = &H400
    Public Const CB_FINDSTRING = WM_USER + 12
    Public Const CB_FINDSTRINGEXACT = WM_USER + 24
    Public Const FB_FINDSTRING = WM_USER + 16
    Public Const FB_FINDSTRINGEXACT = WM_USER + 35
#End If

Public Function FindFirstMatch(ByVal ctlSearch As Control, ByVal SearchString As String, ByVal FirstRow As Integer, ByVal Exact As Boolean) As Integer

#If Win32 Then
    Dim Index As Long
#Else
    Dim Index As Integer
#End If

On Error Resume Next
If TypeOf ctlSearch Is ComboBox Then
    If Exact Then
        Index = SendMessage(ctlSearch.hWnd, CB_FINDSTRINGEXACT, FirstRow, ByVal SearchString)
    Else
        Index = SendMessage(ctlSearch.hWnd, CB_FINDSTRING, FirstRow, ByVal SearchString)
    End If
ElseIf TypeOf ctlSearch Is ListView Then
    If Exact Then
        Index = SendMessage(ctlSearch.hWnd, FB_FINDSTRINGEXACT, FirstRow, ByVal SearchString)
    Else
        Index = SendMessage(ctlSearch.hWnd, FB_FINDSTRING, FirstRow, ByVal SearchString)
    End If
End If

FindFirstMatch = Index

End Function
