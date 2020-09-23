Attribute VB_Name = "Module3"
'Module:   Region.BAS
'               To make irregular shaped windows
'Author:    Pheeraphat Sawangphian
'E-Mail:     tooh@asianet.co.th
'URL:       http://www.geocities.com/Hollywood/Lot/6166
'Note:       There is a single API call, SetWindowRgn.
'                This sets the current window to any region you choose.
'                The Windows API also provides for very flexible region manipulation.
'                Ellipse and polygons can be combined using boolean logic to create irregular regions.
'                These can be applied to the window of your choice.

'Option Explicit 'force explicit declaration of all variables
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

'nCombineMode boolean logic
Public Const RGN_AND = 1
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public Const RGN_DIFF = 4
Public Const RGN_COPY = 5
